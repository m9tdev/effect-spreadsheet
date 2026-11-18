import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { packageRels, partRels, resolveName } from "../src/internal/opc.ts"
import { open } from "../src/internal/zip/archive.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsb from "../src/xlsb/index.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { relsXml, type TestPackage, type TestRel, worksheet, xlsbPackage, xlsxPackage } from "./helpers/opc.ts"
import { failureOf } from "./helpers/run.ts"
import { buildZip } from "./helpers/zip.ts"

const limits = SpreadsheetLimits.defaults

const openArchive = async (entries: ReadonlyArray<{ name: string; data: Uint8Array }>) => {
  const archive = open(await buildZip(entries), limits)
  if ("_tag" in archive) throw new Error(archive._tag)
  return archive
}

describe("opc: resolving relationship targets", () => {
  it("resolves like OPC part names", () => {
    const xl = ["xl"]
    expect(resolveName(xl, "worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml")
    expect(resolveName(xl, "/xl/worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml")
    expect(resolveName(xl, "../docProps/app.xml")).toBe("docProps/app.xml")
    expect(resolveName(xl, "./a%20b.xml")).toBe("xl/a b.xml")
    expect(resolveName(xl, "a//./b.xml")).toBe("xl/a/b.xml")
    // a backslash is an ordinary character in a part name
    expect(resolveName(xl, "worksheets\\sheet1.xml")).toBe("xl/worksheets\\sheet1.xml")
    expect(resolveName([], "xl/workbook.xml")).toBe("xl/workbook.xml")
    expect(resolveName([], "/xl/workbook.xml")).toBe("xl/workbook.xml")
  })

  it("says why a target names no part", () => {
    expect(resolveName(["xl"], "../../escape.xml")).toEqual({ _tag: "Invalid", reason: `"../../escape.xml" climbs above the package root` })
    expect(resolveName([], "../workbook.xml")).toEqual({ _tag: "Invalid", reason: `"../workbook.xml" climbs above the package root` })
    expect(resolveName(["xl"], "%E0%A4%A.xml")).toEqual({ _tag: "Invalid", reason: `"%E0%A4%A.xml" has a bad %-escape` })
    expect(resolveName(["xl"], "a%zz.xml")).toEqual({ _tag: "Invalid", reason: `"a%zz.xml" has a bad %-escape` })
    expect(resolveName(["xl"], "..")).toEqual({ _tag: "Invalid", reason: `".." names no part` })
    expect(resolveName([], "")).toEqual({ _tag: "Invalid", reason: `"" names no part` })
  })
})

describe("opc: package and part relationships", () => {
  const rels: ReadonlyArray<TestRel> = [
    { id: "rPart", kind: "worksheet", target: "worksheets/sheet1.xml" },
    { id: "rMissing", kind: "worksheet", target: "worksheets/nope.xml" },
    { id: "rExternal", kind: "worksheet", target: "worksheets/sheet1.xml", external: true },
    { id: "rClimb", kind: "worksheet", target: "../../x.xml" },
    { id: "rEscape", kind: "worksheet", target: "%E0%A4%A.xml" },
    { id: "rAbsolute", kind: "worksheet", target: "/XL/Worksheets/Sheet1.xml" }
  ]
  const entries = [
    { name: "xl/workbook.xml", data: new Uint8Array() },
    { name: "xl/_rels/workbook.xml.rels", data: relsXml(rels) },
    { name: "xl/worksheets/sheet1.xml", data: new Uint8Array() }
  ]

  it("gives each target its case", async () => {
    const archive = await openArchive(entries)
    const sheet1 = archive.find("xl/worksheets/sheet1.xml")!
    const r = await Effect.runPromise(partRels(archive, archive.find("xl/workbook.xml")!, limits))
    expect(r.byId("rPart")).toEqual({ id: "rPart", kind: "worksheet", target: { _tag: "Part", entry: sheet1 } })
    const part = r.byId("rPart")!.target
    expect(part._tag === "Part" && part.entry === sheet1).toBe(true)
    expect(r.byId("rMissing")!.target).toEqual({ _tag: "Missing", name: "xl/worksheets/nope.xml" })
    expect(r.byId("rExternal")!.target).toEqual({ _tag: "External" })
    expect(r.byId("rClimb")!.target).toEqual({ _tag: "Invalid", reason: `"../../x.xml" climbs above the package root` })
    expect(r.byId("rEscape")!.target).toEqual({ _tag: "Invalid", reason: `"%E0%A4%A.xml" has a bad %-escape` })
    // absolute, and matched case-insensitively: the same entry
    expect(r.byId("rAbsolute")!.target).toEqual({ _tag: "Part", entry: sheet1 })
    expect(r.byId("nope")).toBeUndefined()
  })

  it("byId takes the last of a repeated id; part(kind) the first relationship that names a part", async () => {
    const archive = await openArchive([
      { name: "xl/workbook.xml", data: new Uint8Array() },
      {
        name: "xl/_rels/workbook.xml.rels",
        data: relsXml([
          { id: "r1", kind: "styles", target: "styles.xml", external: true },
          { id: "r1", kind: "styles", target: "../../x.xml" },
          { id: "r1", kind: "styles", target: "missing.xml" },
          { id: "r2", kind: "styles", target: "styles.xml" },
          { id: "r3", kind: "sharedStrings", target: "strings.xml" }
        ])
      },
      { name: "xl/styles.xml", data: new Uint8Array() },
      { name: "xl/strings.xml", data: new Uint8Array() }
    ])
    const r = await Effect.runPromise(partRels(archive, archive.find("xl/workbook.xml")!, limits))
    expect(r.byId("r1")!.target).toEqual({ _tag: "Missing", name: "xl/missing.xml" })
    // external and invalid targets are passed over; a missing one is not, so no styles part
    expect(r.named("styles")).toEqual({ _tag: "Missing", name: "xl/missing.xml" })
    expect(r.part("styles")).toBeUndefined()
    expect(r.part("sharedStrings")).toBe(archive.find("xl/strings.xml"))
    expect(r.named("worksheet")).toBeUndefined()
  })

  it("resolves package relationships from the root; no rels part means no relationships", async () => {
    const archive = await openArchive([
      { name: "_rels/.rels", data: relsXml([{ id: "r", kind: "officeDocument", target: "xl/workbook.xml" }]) },
      { name: "xl/workbook.xml", data: new Uint8Array() }
    ])
    const r = await Effect.runPromise(packageRels(archive, limits))
    expect(r.part("officeDocument")).toBe(archive.find("xl/workbook.xml"))
    const none = await Effect.runPromise(partRels(archive, archive.find("xl/workbook.xml")!, limits))
    expect([none.byId("r"), none.named("worksheet"), none.part("worksheet")]).toEqual([undefined, undefined, undefined])
  })
})

/** Opens the package and reads every sheet; per sheet, its rows' first value or its error reason. */
const readAll = async (format: "xlsx" | "xlsb", p: TestPackage) => {
  const bytes = await (format === "xlsx" ? xlsxPackage(p) : xlsbPackage(p))
  const read = format === "xlsx" ? Xlsx.read : Xlsb.read
  return Effect.runPromise(Effect.exit(Effect.gen(function*() {
    const wb = yield* read(bytes)
    const out: Array<unknown> = []
    for (const sheet of wb.sheets) {
      const rows = yield* Effect.exit(Stream.runCollect(Workbook.rows(wb, sheet.index)))
      out.push(rows._tag === "Success" ? Array.from(rows.value).map((r) => r.cells.length) : (failureOf(rows) as { reason: string }).reason)
    }
    return out
  })))
}

describe.each(["xlsx", "xlsb"] as const)("%s: relationship targets of sheets", (format) => {
  const workbookPath = format === "xlsx" ? "xl/workbook.xml" : "xl/workbook.bin"
  const ext = format === "xlsx" ? "xml" : "bin"
  const office: ReadonlyArray<TestRel> = [{ id: "rId1", kind: "officeDocument", target: workbookPath }]
  const sheetPart = { name: `xl/worksheets/sheet1.${ext}`, data: worksheet(format, 1) }

  it("fails a sheet's read with what is wrong with its target", async () => {
    const targets: ReadonlyArray<TestRel> = [
      { id: "a", kind: "worksheet", target: `worksheets/sheet1.${ext}` },
      { id: "b", kind: "worksheet", target: `worksheets/missing.${ext}` },
      { id: "c", kind: "worksheet", target: "https://example.com/sheet", external: true },
      { id: "d", kind: "worksheet", target: `../../sheet.${ext}` },
      { id: "e", kind: "worksheet", target: `worksheets/%E0%A4%A.${ext}` },
      { id: "f", kind: "chartsheet", target: "https://example.com/chart", external: true },
      { id: "g", kind: "chartsheet", target: "chartsheets/sheet1.xml" }
    ]
    const sheets = [...targets.map((t) => ({ name: t.id, relationshipId: t.id })), { name: "h", relationshipId: "rNone" }]
    const exit = await readAll(format, { sheets, workbookPath, packageRels: office, workbookRels: targets, parts: [sheetPart] })
    expect(exit).toMatchObject({
      _tag: "Success",
      value: [
        [1],
        `sheet part xl/worksheets/missing.${ext} is missing`,
        `sheet "c" has an external target`,
        `sheet "d" has a bad target: "../../sheet.${ext}" climbs above the package root`,
        `sheet "e" has a bad target: "worksheets/%E0%A4%A.${ext}" has a bad %-escape`,
        // a target's error comes before the sheet kind is looked at
        `sheet "f" has an external target`,
        [],
        `sheet "h" has no part`
      ]
    })
  })

  it("finds the workbook part through the package relationships, or at its usual name", async () => {
    const sheets = [{ name: "S", relationshipId: "a" }]
    const workbookRels: ReadonlyArray<TestRel> = [{ id: "a", kind: "worksheet", target: `worksheets/sheet1.${ext}` }]
    const parts = [sheetPart]
    // external and invalid officeDocument targets are passed over: the usual name is used
    const passedOver: ReadonlyArray<TestRel> = [
      { id: "x", kind: "officeDocument", target: "https://example.com/book", external: true },
      { id: "y", kind: "officeDocument", target: "../book" }
    ]
    expect(await readAll(format, { sheets, workbookPath, packageRels: passedOver, workbookRels, parts })).toMatchObject({ _tag: "Success", value: [[1]] })
    expect(await readAll(format, { sheets, workbookPath, workbookRels, parts })).toMatchObject({ _tag: "Success", value: [[1]] })
    // a missing one fails at open, with its name
    const missing = await readAll(format, { sheets, workbookPath, packageRels: [{ id: "z", kind: "officeDocument", target: "/book/main.part" }], workbookRels, parts })
    expect(failureOf(missing)).toMatchObject({ _tag: "MalformedFile", reason: "workbook part book/main.part is missing" })
    // elsewhere, with sheet targets relative to it
    const moved = `book/inner/workbook.${ext}`
    const movedRels: ReadonlyArray<TestRel> = [{ id: "a", kind: "worksheet", target: `../../xl/worksheets/sheet1.${ext}` }]
    expect(await readAll(format, { sheets, workbookPath: moved, packageRels: [{ id: "r", kind: "officeDocument", target: moved }], workbookRels: movedRels, parts }))
      .toMatchObject({ _tag: "Success", value: [[1]] })
  })

  it("sheets may not share a part, present or missing", async () => {
    const twoSheets = [{ name: "A", relationshipId: "a" }, { name: "B", relationshipId: "b" }]
    const present: ReadonlyArray<TestRel> = [
      { id: "a", kind: "worksheet", target: `worksheets/sheet1.${ext}` },
      { id: "b", kind: "worksheet", target: `/XL/worksheets/SHEET1.${ext}` }
    ]
    const shared = await readAll(format, { sheets: twoSheets, workbookPath, packageRels: office, workbookRels: present, parts: [sheetPart] })
    expect(failureOf(shared)).toMatchObject({ reason: `sheets "A" and "B" share part xl/worksheets/sheet1.${ext}` })
    const missing: ReadonlyArray<TestRel> = [
      { id: "a", kind: "worksheet", target: "worksheets/gone.bin" },
      { id: "b", kind: "worksheet", target: "WORKSHEETS/GONE.BIN" }
    ]
    const sharedMissing = await readAll(format, { sheets: twoSheets, workbookPath, packageRels: office, workbookRels: missing })
    expect(failureOf(sharedMissing)).toMatchObject({ reason: `sheets "A" and "B" share part xl/WORKSHEETS/GONE.BIN` })
    // external and invalid targets own nothing
    const bad: ReadonlyArray<TestRel> = [
      { id: "a", kind: "worksheet", target: "same", external: true },
      { id: "b", kind: "worksheet", target: "same", external: true }
    ]
    expect(await readAll(format, { sheets: twoSheets, workbookPath, packageRels: office, workbookRels: bad }))
      .toMatchObject({ _tag: "Success", value: [`sheet "A" has an external target`, `sheet "B" has an external target`] })
  })
})

describe("xlsb: external books by relationship target", () => {
  it("reads only an external book whose link part is present; the others give #REF!", async () => {
    const B = await import("./helpers/xlsb.ts")
    const link = B.concat([B.record(0x168, [...B.u16(0), ...B.wide("rId1"), ...B.u32(0xFFFFFFFF)]), B.supTabs(["Rates"])])
    const books = ["x1", "x2", "x3", "x4", "x5"]
    const globals = [
      B.supSelf(),
      ...books.map((id) => B.supBookSrc(id)),
      B.externSheet(books.map((_, k) => [k + 1, 0, 0] as const)),
      ...books.map((_, k) => B.name(`N${k}`, B.ref3d(k, 0, 0)))
    ]
    const workbookRels: ReadonlyArray<TestRel> = [
      { id: "a", kind: "worksheet", target: "worksheets/sheet1.bin" },
      { id: "x1", kind: "externalLink", target: "externalLinks/externalLink1.bin" },
      { id: "x2", kind: "externalLink", target: "externalLinks/missing.bin" },
      { id: "x3", kind: "externalLink", target: "externalLinks/externalLink1.bin", external: true },
      { id: "x4", kind: "externalLink", target: "../../externalLink1.bin" },
      // the same part as x1, by another spelling: read once, used by both
      { id: "x5", kind: "externalLink", target: "/XL/externalLinks/externalLink1.bin" }
    ]
    const bytes = await xlsbPackage({
      sheets: [{ name: "S", relationshipId: "a" }],
      workbookPath: "xl/workbook.bin",
      packageRels: [{ id: "r", kind: "officeDocument", target: "xl/workbook.bin" }],
      workbookRels,
      parts: [{ name: "xl/worksheets/sheet1.bin", data: worksheet("xlsb", 1) }, { name: "xl/externalLinks/externalLink1.bin", data: link }],
      globals
    })
    const wb = await Effect.runPromise(Xlsb.read(bytes))
    expect(wb.definedNames.map((n) => n.formula)).toEqual(["[1]Rates!$A$1", "#REF!$A$1", "#REF!$A$1", "#REF!$A$1", "[5]Rates!$A$1"])
  })

  it("reads each external link part for its own books", async () => {
    const B = await import("./helpers/xlsb.ts")
    const link = (sheet: string) => B.concat([B.record(0x168, [...B.u16(0), ...B.wide("rId1"), ...B.u32(0xFFFFFFFF)]), B.supTabs([sheet])])
    const books = ["x1", "x2", "x3"]
    const globals = [
      B.supSelf(),
      ...books.map((id) => B.supBookSrc(id)),
      B.externSheet(books.map((_, k) => [k + 1, 0, 0] as const)),
      ...books.map((_, k) => B.name(`N${k}`, B.ref3d(k, 0, 0)))
    ]
    const bytes = await xlsbPackage({
      sheets: [{ name: "S", relationshipId: "a" }],
      workbookPath: "xl/workbook.bin",
      packageRels: [{ id: "r", kind: "officeDocument", target: "xl/workbook.bin" }],
      workbookRels: [
        { id: "a", kind: "worksheet", target: "worksheets/sheet1.bin" },
        { id: "x1", kind: "externalLink", target: "externalLinks/externalLink1.bin" },
        { id: "x2", kind: "externalLink", target: "externalLinks/externalLink2.bin" },
        { id: "x3", kind: "externalLink", target: "externalLinks/externalLink1.bin" }
      ],
      parts: [
        { name: "xl/worksheets/sheet1.bin", data: worksheet("xlsb", 1) },
        { name: "xl/externalLinks/externalLink1.bin", data: link("Rates") },
        { name: "xl/externalLinks/externalLink2.bin", data: link("Prices") }
      ],
      globals
    })
    const wb = await Effect.runPromise(Xlsb.read(bytes))
    expect(wb.definedNames.map((n) => n.formula)).toEqual(["[1]Rates!$A$1", "[2]Prices!$A$1", "[3]Rates!$A$1"])
  })
})

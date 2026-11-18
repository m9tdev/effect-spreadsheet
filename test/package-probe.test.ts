/**
 * How zips are recognised as XLSX and XLSB packages: by the package
 * relationships, as the readers find the workbook part, with the fixed part
 * names as the fallback. Detection and `Spreadsheet.read` follow the same
 * rule, and the archive is opened once per read.
 */
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsb from "../src/xlsb/index.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { counter, metered } from "./helpers/meter.ts"
import { relsXml, type TestRel, worksheet, xlsbPackage, xlsxPackage } from "./helpers/opc.ts"
import { enc, failureOf } from "./helpers/run.ts"
import { buildZip, makeZip } from "./helpers/zip.ts"

const run = <A, E>(effect: Effect.Effect<A, E>) => Effect.runPromise(Effect.exit(effect))
const formatOf = async (bytes: Uint8Array) => {
  const exit = await run(Spreadsheet.detectFormat(bytes))
  return Exit.isSuccess(exit) ? exit.value : failureOf(exit)
}
/** The first cell's value of the first sheet, or the error. */
const firstCell = async (read: Effect.Effect<Workbook.Workbook, unknown>) => {
  const exit = await run(Effect.flatMap(read, (wb) => Stream.runCollect(Workbook.rows(wb, 0))))
  return Exit.isSuccess(exit) ? [...exit.value][0]?.cells[0]?.value : failureOf(exit)
}

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
const XLSB_TYPE = "application/vnd.ms-excel.sheet.binary.macroEnabled.main"
const contentTypes = (defaults: ReadonlyArray<readonly [string, string]>, overrides: ReadonlyArray<readonly [string, string]>) => ({
  name: "[Content_Types].xml",
  data: enc(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${
      defaults.map(([extension, type]) => `<Default Extension="${extension}" ContentType="${type}"/>`).join("")
    }${overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join("")}</Types>`
  )
})
const officeDocument = (target: string): ReadonlyArray<TestRel> => [{ id: "rId1", kind: "officeDocument", target }]

/** A one-sheet package whose workbook part is `workbookPath`, its sheet next to it. */
const book = (
  format: "xlsx" | "xlsb",
  workbookPath: string,
  options: { readonly packageRels?: ReadonlyArray<TestRel>; readonly parts?: ReadonlyArray<{ name: string; data: Uint8Array }> } = {}
) => {
  const dir = workbookPath.slice(0, workbookPath.lastIndexOf("/") + 1)
  const sheet = format === "xlsx" ? "sheet1.xml" : "sheet1.bin"
  return (format === "xlsx" ? xlsxPackage : xlsbPackage)({
    sheets: [{ name: "S", relationshipId: "r1" }],
    workbookPath,
    packageRels: options.packageRels ?? officeDocument(workbookPath),
    workbookRels: [{ id: "r1", kind: "worksheet", target: sheet }],
    parts: [{ name: dir + sheet, data: worksheet(format, 42) }, ...(options.parts ?? [])]
  })
}

describe("zip packages: the package relationships decide", () => {
  it("an XLSX with a stray xl/workbook.bin is xlsx, and reads", async () => {
    const bytes = await book("xlsx", "xl/workbook.xml", { parts: [{ name: "xl/workbook.bin", data: new Uint8Array([1, 2, 3]) }] })
    expect(await formatOf(bytes)).toBe("xlsx")
    expect(await firstCell(Spreadsheet.read(bytes))).toBe(42)
    expect(await firstCell(Xlsx.read(bytes))).toBe(42)
  })

  it("an XLSB with a stray xl/workbook.xml is xlsb, and reads", async () => {
    const bytes = await book("xlsb", "xl/workbook.bin", { parts: [{ name: "xl/workbook.xml", data: enc("<workbook/>") }] })
    expect(await formatOf(bytes)).toBe("xlsb")
    expect(await firstCell(Spreadsheet.read(bytes))).toBe(42)
  })

  it("a workbook part elsewhere goes by its content type: an Override, else the Default for its extension", async () => {
    const xlsx = await book("xlsx", "book/wb.xml", {
      parts: [contentTypes([["xml", "application/xml"]], [["/book/wb.xml", XLSX_TYPE]])]
    })
    expect(await formatOf(xlsx)).toBe("xlsx")
    expect(await firstCell(Spreadsheet.read(xlsx))).toBe(42)
    expect(await firstCell(Xlsx.read(xlsx))).toBe(42)
    const xlsb = await book("xlsb", "book/wb.bin", { parts: [contentTypes([["bin", XLSB_TYPE]], [])] })
    expect(await formatOf(xlsb)).toBe("xlsb")
    expect(await firstCell(Spreadsheet.read(xlsb))).toBe(42)
    expect(await firstCell(Xlsb.read(xlsb))).toBe(42)
    // an Override wins over the Default for its extension
    const both = await book("xlsx", "book/wb.xml", {
      parts: [contentTypes([["xml", XLSB_TYPE]], [["/book/wb.xml", XLSX_TYPE]])]
    })
    expect(await formatOf(both)).toBe("xlsx")
  })

  it("content types and part names match case-insensitively, part names with %-escapes decoded", async () => {
    const bytes = await book("xlsx", "Book Data/WB.xml", {
      packageRels: officeDocument("Book%20Data/WB.xml"),
      parts: [contentTypes([], [["/book%20data/wb.XML", XLSX_TYPE.toUpperCase()]])]
    })
    expect(await formatOf(bytes)).toBe("xlsx")
    expect(await firstCell(Spreadsheet.read(bytes))).toBe(42)
    const byExtension = await book("xlsb", "book/wb.BIN", { parts: [contentTypes([["Bin", `${XLSB_TYPE}; x=1`]], [])] })
    expect(await formatOf(byExtension)).toBe("xlsb")
  })

  it("the macro-enabled and template workbook content types are xlsx", async () => {
    for (
      const type of [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        "application/vnd.ms-excel.template.macroEnabled.main+xml",
        "application/vnd.ms-excel.addin.macroEnabled.main+xml"
      ]
    ) {
      const bytes = await book("xlsx", "book/wb.xml", { parts: [contentTypes([], [["/book/wb.xml", type]])] })
      expect(await formatOf(bytes), type).toBe("xlsx")
    }
  })

  it("a package whose main part isn't a workbook is a zip, even with a workbook part name in it", async () => {
    const docx = (extra: ReadonlyArray<{ name: string; data: Uint8Array }>) =>
      buildZip([
        { name: "_rels/.rels", data: relsXml(officeDocument("word/document.xml")) },
        { name: "word/document.xml", data: enc("<document/>") },
        contentTypes([["xml", "application/xml"]], [[
          "/word/document.xml",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        ]]),
        ...extra
      ])
    for (const bytes of [await docx([]), await docx([{ name: "xl/workbook.xml", data: enc("<workbook/>") }])]) {
      expect(await formatOf(bytes)).toBe("zip")
      expect(failureOf(await run(Spreadsheet.read(bytes)))).toMatchObject({ _tag: "UnsupportedFormat", format: "zip" })
    }
    // no content type at all: no content types part, or none for the part
    const bare = await book("xlsx", "book/wb.xml")
    expect(await formatOf(bare)).toBe("zip")
    const untyped = await book("xlsx", "book/wb.xml", { parts: [contentTypes([["bin", XLSB_TYPE]], [["/book/other.xml", XLSX_TYPE]])] })
    expect(await formatOf(untyped)).toBe("zip")
  })

  it("without an officeDocument relationship, or with a missing target, the fixed names decide", async () => {
    // no _rels/.rels: the binary part wins, as before
    expect(await formatOf(makeZip([{ name: "xl/workbook.xml" }, { name: "xl/workbook.bin" }]))).toBe("xlsb")
    // _rels/.rels without an officeDocument relationship
    const noMain = await book("xlsx", "xl/workbook.xml", { packageRels: [{ id: "rId1", kind: "core-properties", target: "docProps/core.xml" }] })
    expect(await formatOf(noMain)).toBe("xlsx")
    expect(await firstCell(Spreadsheet.read(noMain))).toBe(42)
    // a target the package lacks: detected by the fixed name, and reading reports the missing part, as Xlsx.read does
    const missing = await book("xlsx", "xl/workbook.xml", { packageRels: officeDocument("xl/gone.xml") })
    expect(await formatOf(missing)).toBe("xlsx")
    const expected = { _tag: "MalformedFile", format: "xlsx", reason: "workbook part xl/gone.xml is missing" }
    expect(await firstCell(Spreadsheet.read(missing))).toMatchObject(expected)
    expect(await firstCell(Xlsx.read(missing))).toMatchObject(expected)
    // a missing target and no fixed name: just a zip
    const nothing = await buildZip([{ name: "_rels/.rels", data: relsXml(officeDocument("xl/gone.xml")) }])
    expect(await formatOf(nothing)).toBe("zip")
  })
})

describe("zip packages: broken ones fail, other zips don't", () => {
  it("a package-looking zip the archive rejects is MalformedFile; another zip stays a zip", async () => {
    for (const marker of ["_rels/.rels", "xl/workbook.xml", "/XL/Workbook.bin"]) {
      const bytes = makeZip([{ name: marker }, { name: "README" }, { name: "readme" }])
      for (const exit of [await run(Spreadsheet.detectFormat(bytes)), await run(Spreadsheet.read(bytes))]) {
        expect(failureOf(exit), marker).toMatchObject({ _tag: "MalformedFile", format: "zip", reason: "duplicate entry name" })
      }
    }
    const plain = makeZip([{ name: "README" }, { name: "readme" }])
    expect(await formatOf(plain)).toBe("zip")
    expect(failureOf(await run(Spreadsheet.read(plain)))).toMatchObject({ _tag: "UnsupportedFormat", format: "zip" })
  })

  it("package relationships that don't read fail detection as they fail reading", async () => {
    // makeZip writes no CRC-32: the relationships part fails its check
    const bytes = makeZip([{ name: "_rels/.rels", data: relsXml(officeDocument("xl/workbook.xml")) }, { name: "xl/workbook.xml" }])
    const expected = { _tag: "MalformedFile", format: "zip", reason: "CRC-32 mismatch", part: "_rels/.rels" }
    expect(await formatOf(bytes)).toMatchObject(expected)
    expect(failureOf(await run(Spreadsheet.read(bytes)))).toMatchObject(expected)
    expect(failureOf(await run(Xlsx.read(bytes)))).toMatchObject(expected)
  })

  it("a content types part that doesn't read fails detection of a workbook part elsewhere", async () => {
    const bytes = await book("xlsx", "book/wb.xml", { parts: [{ name: "[Content_Types].xml", data: enc("<Types><Override") }] })
    expect(await formatOf(bytes)).toMatchObject({ _tag: "MalformedFile", part: "[Content_Types].xml" })
  })

  it("an ODF mimetype is read with the archive's checks on its entry", async () => {
    const ods = enc("application/vnd.oasis.opendocument.spreadsheet")
    const good = makeZip([{ name: "mimetype", data: ods }, { name: "content.xml" }])
    expect(await formatOf(good)).toBe("ods")
    // the local header names another entry: not read, so just a zip
    const renamed = good.slice()
    renamed.set(enc("nimetype"), 30)
    expect(await formatOf(renamed)).toBe("zip")
    // a deflated mimetype isn't read either
    expect(await formatOf(await buildZip([{ name: "mimetype", data: ods, method: "deflate" }]))).toBe("zip")
    expect(await formatOf(await buildZip([{ name: "mimetype", data: ods, method: "store" }]))).toBe("ods")
  })
})

describe("zip packages: opened once per read", () => {
  it("Spreadsheet.read inflates the package relationships once, as Xlsx.read and Xlsb.read do", async () => {
    const inflated = async (read: Effect.Effect<unknown, unknown>) => {
      const meter = counter("bytes")
      await Effect.runPromise(Effect.ignore(metered(meter, read)))
      return meter.total()
    }
    for (const format of ["xlsx", "xlsb"] as const) {
      const path = format === "xlsx" ? "xl/workbook.xml" : "xl/workbook.bin"
      const bytes = await book(format, path)
      const reader = await inflated(format === "xlsx" ? Xlsx.read(bytes) : Xlsb.read(bytes))
      expect(reader).toBeGreaterThan(0)
      expect(await inflated(Spreadsheet.read(bytes)), format).toBe(reader)
      // detection inflates the package relationships only
      expect(await inflated(Spreadsheet.detectFormat(bytes)), format).toBe(relsXml(officeDocument(path)).length)
    }
  })
})

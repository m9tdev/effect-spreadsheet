import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { formatFor, isDateCode, serialToDate } from "../src/internal/xlsx/numFmt.ts"
import { resolvePart } from "../src/internal/xlsx/parts.ts"
import { columnName, parseCellRef, parseRange } from "../src/internal/xlsx/refs.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { failureOf } from "./helpers/run.ts"
import { sheetXml, type TestWorkbook, xlsx } from "./helpers/xlsx.ts"
import { buildZip } from "./helpers/zip.ts"

const enc = (s: string) => new TextEncoder().encode(s)

/** Reads a workbook and every sheet's rows; the exit. */
const readExit = async (bytes: Uint8Array, limits?: LimitOverrides) =>
  Effect.runPromise(Effect.exit(Effect.gen(function*() {
    const wb = yield* Xlsx.read(bytes, { limits })
    const sheets: Array<Array<Workbook.Row>> = []
    for (const sheet of wb.sheets) sheets.push(Array.from(yield* Stream.runCollect(Workbook.rows(wb, sheet.index))))
    return { wb, sheets }
  })))

const readOk = async (book: TestWorkbook, limits?: LimitOverrides) => {
  const exit = await readExit(await xlsx(book), limits)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value
}

const errorOf = async (book: TestWorkbook, limits?: LimitOverrides) => failureOf(await readExit(await xlsx(book), limits))

const values = (rows: ReadonlyArray<Workbook.Row>) =>
  rows.map((r) => [r.index, r.hidden, r.cells.map((c) => [c.column, c._tag, Cell.value(c) instanceof Date ? (Cell.value(c) as Date).toISOString() : Cell.value(c)])])

describe("number formats and serial dates", () => {
  it("tells date/time codes from number codes", () => {
    for (const code of ["yyyy-mm-dd", "m/d/yy", "h:mm", "[h]:mm:ss", "[mm]:ss", "mmss.0", "[$-409]mmmm d, yyyy;@", "dd/mm/yyyy hh:mm:ss.000", "AM/PM h", "\"Date:\" yyyy", "yyyy\\-mm"]) {
      expect(isDateCode(code), code).toBe(true)
    }
    for (const code of ["General", "0.00", "#,##0", "0.00E+00", "@", "[Red][<0]\"neg\";0", "\"days\" 0", "0 \\h", "_(* #,##0_)", "# ?/?", "[Color10]0", "0%"]) {
      expect(isDateCode(code), code).toBe(false)
    }
  })

  it("built-in format ids, custom codes winning, locale date ids", () => {
    expect(formatFor(0, new Map())).toEqual({ code: undefined, date: false })
    expect(formatFor(14, new Map())).toEqual({ code: "mm-dd-yy", date: true })
    expect(formatFor(4, new Map())).toEqual({ code: "#,##0.00", date: false })
    expect(formatFor(14, new Map([[14, "0.00"]]))).toEqual({ code: "0.00", date: false })
    expect(formatFor(31, new Map())).toEqual({ code: undefined, date: true })
    expect(formatFor(500, new Map())).toEqual({ code: undefined, date: false })
  })

  it("1900 system: 1 is 1900-01-01, the fictional day 60 has no date, 61 is 1900-03-01; 1904 system: 0 is 1904-01-01", () => {
    const iso = (s: number, d1904 = false) => serialToDate(s, d1904)?.toISOString()
    expect(iso(0)).toBe("1899-12-31T00:00:00.000Z")
    expect(iso(1)).toBe("1900-01-01T00:00:00.000Z")
    expect(iso(59.5)).toBe("1900-02-28T12:00:00.000Z")
    expect(iso(60)).toBeUndefined()
    expect(iso(60.9)).toBeUndefined()
    expect(iso(61)).toBe("1900-03-01T00:00:00.000Z")
    expect(iso(46024.5208333333)).toBe("2026-01-02T12:30:00.000Z") // rounded to the millisecond
    expect(iso(2958465.5)).toBe("9999-12-31T12:00:00.000Z")
    expect(iso(2958466)).toBeUndefined()
    expect(iso(-1)).toBeUndefined()
    expect(iso(0, true)).toBe("1904-01-01T00:00:00.000Z")
    expect(iso(46024.5208333333, true)).toBe("2030-01-03T12:30:00.000Z")
  })
})

describe("references and part names", () => {
  it("parses A1 references and ranges", () => {
    expect(parseCellRef("A1")).toEqual({ column: 0, row: 0 })
    expect(parseCellRef("$XFD$1048576")).toEqual({ column: 16383, row: 1048575 })
    expect(parseCellRef("aa10")).toEqual({ column: 26, row: 9 })
    expect(parseCellRef("C")).toEqual({ column: 2, row: undefined })
    expect(parseCellRef("ABCDEFGH1")?.column).toBe(Infinity)
    expect(parseCellRef("A12345678901")?.row).toBe(Infinity)
    for (const bad of ["", "A0", "1A", "A-1", "A1:B2", "Ä1"]) expect(parseCellRef(bad), bad).toBeUndefined()
    expect(parseRange("B2:A1")).toEqual({ firstRow: 0, firstColumn: 0, lastRow: 1, lastColumn: 1 })
    expect(parseRange("C3")).toEqual({ firstRow: 2, firstColumn: 2, lastRow: 2, lastColumn: 2 })
    expect(parseRange("A1:")).toBeUndefined()
    expect([0, 25, 26, 701, 702, 16383].map(columnName)).toEqual(["A", "Z", "AA", "ZZ", "AAA", "XFD"])
  })

  it("resolves relationship targets like OPC part names", () => {
    expect(resolvePart("xl/workbook.xml", "worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml")
    expect(resolvePart("xl/workbook.xml", "/xl/worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml")
    expect(resolvePart("xl/workbook.xml", "../docProps/app.xml")).toBe("docProps/app.xml")
    expect(resolvePart("xl/workbook.xml", "./a%20b.xml")).toBe("xl/a b.xml")
    expect(resolvePart("xl/workbook.xml", "../../escape.xml")).toBeUndefined()
    expect(resolvePart("xl/workbook.xml", "%E0%A4%A.xml")).toBeUndefined()
  })
})

describe("xlsx: cells and rows", () => {
  it("cell types, formulas, styled blanks and number formats", async () => {
    const { sheets } = await readOk({
      sharedStrings: ["shared", "a_x0041_b"],
      xfs: [0, 14, 164, 4],
      numFmts: { 164: "0.000" },
      sheets: [{
        name: "S",
        data: `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><r><t>in</t></r><r><t>line</t></r><rPh><t>skip</t></rPh></is></c>` +
          `<c r="D1" t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c r="E1" t="b"><v>0</v></c><c r="F1" t="e"><v>#REF!</v></c></row>` +
          `<row r="2"><c r="A2" s="1"><v>61.25</v></c><c r="B2" s="2"><v>1.5</v></c><c r="C2" s="3"><v>-0</v></c><c r="D2" s="1"/><c r="E2"><f>NOW()</f></c>` +
          `<c r="F2" t="d"><v>2026-01-02T03:04:05</v></c><c r="G2" t="d"><v>2026-01-02T03:04:05+02:00</v></c><c r="H2" s="1"><v>60</v></c></row>`
      }]
    })
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Text", "shared"], [1, "Text", "aAb"], [2, "Text", "inline"], [3, "Text", "xy"], [4, "Boolean", false], [5, "Error", "#REF!"]]],
      [1, false, [
        [0, "Date", "1900-03-01T06:00:00.000Z"],
        [1, "Number", 1.5],
        [2, "Number", 0],
        [3, "Empty", undefined],
        [4, "Empty", undefined],
        [5, "Date", "2026-01-02T03:04:05.000Z"],
        [6, "Date", "2026-01-02T01:04:05.000Z"],
        [7, "Number", 60]
      ]]
    ])
    const row2 = sheets[0]![1]!.cells
    expect(row2.map((c) => c.numFmt)).toEqual(["mm-dd-yy", "0.000", "#,##0.00", undefined, undefined, undefined, undefined, "mm-dd-yy"])
    expect(sheets[0]![0]!.cells[3]!.formula).toBe('"x"&"y"')
    expect(row2[4]!.formula).toBe("NOW()")
  })

  it("number formats come only from <numFmts>, not from conditional formats (<dxfs>)", async () => {
    const { sheets } = await readOk({
      xfs: [0, 164],
      numFmts: { 164: "0.00" },
      stylesExtra: `<dxfs count="1"><dxf><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></dxf></dxfs>`,
      sheets: [{ name: "S", data: `<row><c s="1"><v>45000</v></c></row>` }]
    })
    expect(sheets[0]![0]!.cells[0]).toEqual(Cell.number(0, 45000, { numFmt: "0.00" }))
  })

  it("rows and cells without references, out-of-order cells, hidden rows", async () => {
    const { sheets } = await readOk({
      sheets: [{
        name: "S",
        data: `<row><c><v>1</v></c><c><v>2</v></c></row><row hidden="true"/><row r="7"><c r="C7"><v>3</v></c><c><v>4</v></c></row>` +
          `<row><c r="B9"><v>5</v></c></row><row r="12"><c r="C12"><v>6</v></c><c r="A12"><v>7</v></c><c r="C12"><v>8</v></c></row>`
      }]
    })
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Number", 1], [1, "Number", 2]]],
      [1, true, []],
      [6, false, [[2, "Number", 3], [3, "Number", 4]]],
      [8, false, [[1, "Number", 5]]],
      [11, false, [[0, "Number", 7], [2, "Number", 8]]]
    ])
  })

  it("the 1904 date system", async () => {
    const { sheets } = await readOk({ date1904: true, xfs: [0, 14], sheets: [{ name: "S", data: `<row><c s="1"><v>0</v></c><c s="1"><v>60</v></c></row>` }] })
    expect(values(sheets[0]!)[0]![2]).toEqual([[0, "Date", "1904-01-01T00:00:00.000Z"], [1, "Date", "1904-03-01T00:00:00.000Z"]])
  })

  it("sheet visibility, defined names, chart sheets, rows re-runnable", async () => {
    const bytes = await xlsx({
      sheets: [{ name: "A", data: `<row><c><v>1</v></c></row>` }, { name: "B", state: "hidden" }, { name: "C", state: "veryHidden", kind: "chartsheet" }],
      definedNames: `<definedName name="Total">A!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="1" hidden="1">B!$A$1:$C$3</definedName><definedName name="Bad" localSheetId="7">#REF!</definedName>`
    })
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    expect(wb.sheets).toEqual([
      { name: "A", index: 0, visibility: "visible" },
      { name: "B", index: 1, visibility: "hidden" },
      { name: "C", index: 2, visibility: "veryHidden" }
    ])
    expect(wb.definedNames).toEqual([
      { name: "Total", formula: "A!$A$1", hidden: false },
      { name: "_xlnm.Print_Area", formula: "B!$A$1:$C$3", hidden: true, scope: 1 },
      { name: "Bad", formula: "#REF!", hidden: false }
    ])
    const rows = Workbook.rows(wb, "a")
    expect(values(Array.from(await Effect.runPromise(Stream.runCollect(rows))))).toEqual([[0, false, [[0, "Number", 1]]]])
    expect(Array.from(await Effect.runPromise(Stream.runCollect(rows))).length).toBe(1)
    expect(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, "C"))))).toEqual([])
  })

  it("finds the workbook through the relationships, wherever it is", async () => {
    const { wb, sheets } = await readOk({ workbookPath: "book/main.xml", sharedStrings: ["moved"], sheets: [{ name: "S", data: `<row><c t="s"><v>0</v></c></row>` }] })
    expect(wb.sheets[0]!.name).toBe("S")
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Text", "moved"]]]])
  })

  it("an xlsm (macro parts included) reads like an xlsx, through Spreadsheet.read", async () => {
    const bytes = await xlsx({ sheets: [{ name: "S", data: `<row><c><v>1</v></c></row>` }], extra: [{ name: "xl/vbaProject.bin", data: new Uint8Array([0xD0, 0xCF, 0x11, 0xE0]) }] })
    const wb = await Effect.runPromise(Spreadsheet.read(bytes))
    expect(wb.format).toBe("xlsx")
    expect(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0)))).length).toBe(1)
  })

  it("limits from the layer at read time stay with the workbook", async () => {
    const bytes = await xlsx({ sheets: [{ name: "S", data: `<row><c><v>1</v></c></row><row><c><v>2</v></c></row><row><c><v>3</v></c></row>` }] })
    const wb = await Effect.runPromise(Xlsx.read(bytes).pipe(Effect.provide(SpreadsheetLimits.layer({ maxRows: 2 }))))
    const error = failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)))))
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 2 })
    expect(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0, { limits: { maxRows: 3 } })))).length).toBe(3)
  })
})

describe("xlsx: sheet details", () => {
  const withLayout = (layout: string, data = `<row><c><v>1</v></c></row>`) =>
    xlsx({ sheets: [{ name: "S", raw: `<worksheet xmlns="x"><dimension ref="B2:D9"/><cols>${layout}</cols><sheetData>${data}</sheetData><mergeCells><mergeCell ref="A1:C1"/><mergeCell ref="B3:B4"/></mergeCells></worksheet>` }] })
  const details = async (bytes: Uint8Array, limits?: LimitOverrides) =>
    Effect.runPromise(Effect.exit(Effect.flatMap(Xlsx.read(bytes, { limits }), (wb) => Workbook.sheetDetails(wb, 0))))

  it("hidden columns as merged spans, merges, the declared dimension", async () => {
    const exit = await details(await withLayout(`<col min="2" max="3" hidden="1"/><col min="4" max="4" hidden="true"/><col min="6" max="6" width="9"/><col min="8" max="20000" hidden="1"/><col min="9" max="9" hidden="1"/>`))
    expect(Exit.isSuccess(exit) && exit.value).toEqual({
      hiddenColumns: [{ first: 1, last: 3 }, { first: 7, last: 16383 }],
      merges: [
        { firstRow: 0, firstColumn: 0, lastRow: 0, lastColumn: 2 },
        { firstRow: 2, firstColumn: 1, lastRow: 3, lastColumn: 1 }
      ],
      dimension: { firstRow: 1, firstColumn: 1, lastRow: 8, lastColumn: 3 }
    })
  })

  it("cell data is skipped, CSV has no details, missing sheets fail", async () => {
    const bytes = await withLayout("", `<row><c><v>1</v></c></row>`.repeat(10_000))
    const exit = await details(bytes)
    expect(Exit.isSuccess(exit) && exit.value.merges.length).toBe(2)
    const csv = await Effect.runPromise(Effect.flatMap(Spreadsheet.read(enc("a,b\n")), (wb) => Workbook.sheetDetails(wb, 0)))
    expect(csv).toEqual({ hiddenColumns: [], merges: [] })
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    expect(failureOf(await Effect.runPromise(Effect.exit(Workbook.sheetDetails(wb, "nope"))))).toMatchObject({ _tag: "SheetNotFound" })
  })

  it("bad ranges are MalformedFile; spans and merges have limits", async () => {
    expect(failureOf(await details(await withLayout(`<col min="0" max="2" hidden="1"/>`)))).toMatchObject({ _tag: "MalformedFile" })
    expect(failureOf(await details(await withLayout(`<col min="x" max="2" hidden="1"/>`)))).toMatchObject({ _tag: "MalformedFile" })
    const badMerge = await xlsx({ sheets: [{ name: "S", raw: `<worksheet xmlns="x"><sheetData/><mergeCells><mergeCell ref="A1:"/></mergeCells></worksheet>` }] })
    expect(failureOf(await details(badMerge))).toMatchObject({ _tag: "MalformedFile" })
    expect(failureOf(await details(await withLayout(`<col min="1" max="1"/>`.repeat(5)), { maxColumns: 4 }))).toMatchObject({ limit: "maxColumns" })
    expect(failureOf(await details(await withLayout(""), { maxCellsPerSheet: 1 }))).toMatchObject({ limit: "maxCellsPerSheet" })
  })
})

describe("xlsx: hostile and broken files", () => {
  const sheet = (data: string): TestWorkbook => ({ sheets: [{ name: "S", data }] })

  it("a huge declared dimension and a far cell cost only the cells present", async () => {
    const { sheets } = await readOk({
      sheets: [{ name: "S", raw: `<worksheet xmlns="x"><dimension ref="A1:XFD1048576"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="1048576"><c r="XFD1048576"><v>2</v></c></row></sheetData></worksheet>` }]
    })
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Number", 1]]], [1048575, false, [[16383, "Number", 2]]]])
  })

  it("row and column indexes past the limits are LimitExceeded", async () => {
    expect(await errorOf(sheet(`<row r="1048577"/>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows" })
    expect(await errorOf(sheet(`<row r="99999999999999999"/>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows" })
    // an implicit row after the last allowed one
    expect(await errorOf(sheet(`<row r="1048576"/><row/>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", actual: 1048577 })
    expect(await errorOf(sheet(`<row><c r="XFE1"><v>1</v></c></row>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxColumns", actual: 16385 })
    expect(await errorOf(sheet(`<row><c r="AAAAAAAAAAAA1"><v>1</v></c></row>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxColumns" })
    expect(await errorOf(sheet(`<row><c r="B5"><v>1</v></c></row>`), { maxColumns: 1 })).toMatchObject({ limit: "maxColumns" })
  })

  it("many rows stop at maxRows, cells at maxCellsPerSheet, after a bounded number of rows", async () => {
    let seen = 0
    const bytes = await xlsx(sheet(`<row/>`.repeat(100_000)))
    const exit = await Effect.runPromise(Effect.exit(Effect.flatMap(Xlsx.read(bytes, { limits: { maxRows: 1000 } }), (wb) =>
      Stream.runForEach(Workbook.rows(wb, 0), () => Effect.sync(() => seen++)))))
    expect(failureOf(exit)).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 1000 })
    expect(seen).toBeLessThanOrEqual(1000)
    expect(await errorOf(sheet(`<row><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>`), { maxCellsPerSheet: 2 })).toMatchObject({ limit: "maxCellsPerSheet" })
    // rows repeating an index still count
    expect(await errorOf(sheet(`<row r="1"/>`.repeat(5)), { maxRows: 3 })).toMatchObject({ limit: "maxRows", actual: 4 })
  })

  it("sheets, shared strings and string lengths have limits", async () => {
    expect(await errorOf({ sheets: [{ name: "A" }, { name: "B" }, { name: "C" }] }, { maxSheets: 2 })).toMatchObject({ limit: "maxSheets", max: 2 })
    expect(await errorOf({ sharedStrings: ["a", "b", "c"], sheets: [{ name: "S" }] }, { maxSharedStrings: 2 })).toMatchObject({ limit: "maxSharedStrings" })
    expect(await errorOf({ sharedStringsXml: `<sst xmlns="x"><si>${"<r><t>abcd</t></r>".repeat(50)}</si></sst>`, sheets: [{ name: "S" }] }, { maxStringLength: 100 })).toMatchObject({ limit: "maxStringLength" })
    expect(await errorOf(sheet(`<row><c><f>${"x".repeat(500)}</f></c></row>`), { maxStringLength: 100 })).toMatchObject({ limit: "maxStringLength" })
  })

  it("deep XML, DTDs and broken XML in a sheet are typed errors", async () => {
    expect(await errorOf(sheet(`<row><c><v>${"<x>".repeat(200)}${"</x>".repeat(200)}</v></c></row>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxXmlDepth" })
    expect(await errorOf({ sheets: [{ name: "S", raw: `<!DOCTYPE x [<!ENTITY a "b">]><worksheet/>` }] })).toMatchObject({ _tag: "MalformedFile" })
    expect(await errorOf(sheet(`<row><c><v>1</c></row>`))).toMatchObject({ _tag: "MalformedFile" })
    expect(await errorOf({ sheets: [{ name: "S", raw: `<worksheet><sheetData><row>` }] })).toMatchObject({ _tag: "MalformedFile" })
  })

  it("bad values are MalformedFile", async () => {
    for (const cell of [`<c t="s"><v>5</v></c>`, `<c t="s"><v>-1</v></c>`, `<c><v>1,5</v></c>`, `<c><v>0x10</v></c>`, `<c><v>1e999</v></c>`, `<c t="b"><v>2</v></c>`, `<c t="d"><v>yesterday</v></c>`, `<c t="zz"><v>1</v></c>`, `<c r="1A"><v>1</v></c>`]) {
      const error = await errorOf({ sharedStrings: ["only"], sheets: [{ name: "S", data: `<row>${cell}</row>` }] })
      expect(error, cell).toMatchObject({ _tag: "MalformedFile", format: "xlsx" })
    }
    expect(await errorOf(sheet(`<row r="x"/>`))).toMatchObject({ _tag: "MalformedFile" })
    expect(await errorOf(sheet(`<row r="0"/>`))).toMatchObject({ _tag: "MalformedFile" })
  })

  it("broken packages: no workbook part, a sheet without a part, a target outside the package", async () => {
    const noWorkbook = await buildZip([{ name: "_rels/.rels", data: enc(`<Relationships xmlns="x"/>`) }, { name: "xl/other.xml", data: enc("<a/>") }])
    expect(failureOf(await readExit(noWorkbook))).toMatchObject({ _tag: "MalformedFile", reason: "workbook part xl/workbook.xml is missing" })
    // a sheet part that isn't there: that sheet fails, the others read
    const missing = await xlsx({
      sheets: [{ name: "Good", data: `<row><c><v>1</v></c></row>` }, { name: "Gone" }],
      omit: ["xl/worksheets/sheet2.xml"]
    })
    const wb = await Effect.runPromise(Xlsx.read(missing))
    expect(wb.sheets.map((s) => s.name)).toEqual(["Good", "Gone"])
    expect(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, "Good")))).length).toBe(1)
    const runs: ReadonlyArray<Effect.Effect<unknown, unknown>> = [Stream.runCollect(Workbook.rows(wb, "Gone")), Workbook.sheetDetails(wb, "Gone")]
    for (const run of runs) {
      expect(failureOf(await Effect.runPromise(Effect.exit(run)))).toMatchObject({
        _tag: "MalformedFile",
        format: "xlsx",
        reason: "sheet part xl/worksheets/sheet2.xml is missing"
      })
    }
    const relsXml = `<Relationships xmlns="x"><Relationship Id="rId1" Type="http://x/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://x/worksheet" Target="../../../etc/passwd"/></Relationships>`
    const two = await buildZip([
      { name: "_rels/.rels", data: enc(`<Relationships xmlns="x"><Relationship Id="r" Type="http://x/officeDocument" Target="xl/workbook.xml"/></Relationships>`) },
      { name: "xl/workbook.xml", data: enc(`<workbook xmlns="x" xmlns:r="y"><sheets><sheet name="Good" r:id="rId1"/><sheet name="Outside" r:id="rId2"/><sheet name="NoRel" r:id="rId9"/></sheets></workbook>`) },
      { name: "xl/_rels/workbook.xml.rels", data: enc(relsXml) },
      { name: "xl/worksheets/sheet1.xml", data: enc(sheetXml(`<row><c><v>1</v></c></row>`)) }
    ])
    const wb2 = await Effect.runPromise(Xlsx.read(two))
    expect(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb2, "Good")))).length).toBe(1)
    for (const name of ["Outside", "NoRel"]) {
      const error = failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb2, name)))))
      expect(error, name).toMatchObject({ _tag: "MalformedFile", format: "xlsx" })
    }
  })

  it("a sheet that inflates past the archive budget stops there", async () => {
    // ~96 MiB of whitespace and unknown elements inside sheetData: tiny
    // compressed, past the 64 MiB floor
    const data = (" ".repeat(1000) + "<x/>").repeat(96 * 1024)
    const error = await errorOf(sheet(data))
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxInflateRatio" })
  }, 60_000)

  it("encrypted entries are Encrypted", async () => {
    const bytes = await buildZip([
      { name: "_rels/.rels", data: enc(`<Relationships xmlns="x"/>`), flags: 0x801 },
      { name: "xl/workbook.xml", data: enc("<workbook/>") }
    ])
    expect(failureOf(await readExit(bytes))).toMatchObject({ _tag: "Encrypted" })
  })
})

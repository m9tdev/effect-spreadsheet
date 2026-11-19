import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { addFormat, cellStyles, classify, isDateCode, type NumberFormat, serialToDate } from "../src/internal/numFmt.ts"
import { columnName, parseCellRef, parseRange } from "../src/internal/refs.ts"
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
    for (const code of ["yyyy-mm-dd", "m/d/yy", "h:mm", "[h]:mm:ss", "[h]", "[ss]", "[mm]:ss", "mmss.0", "[$-409]mmmm d, yyyy;@", "dd/mm/yyyy hh:mm:ss.000", "AM/PM h", "\"Date:\" yyyy", "yyyy\\-mm"]) {
      expect(isDateCode(code), code).toBe(true)
    }
    for (const code of ["General", "0.00", "#,##0", "0.00E+00", "@", "[Red][<0]\"neg\";0", "\"days\" 0", "0 \\h", "_(* #,##0_)", "# ?/?", "[Color10]0", "0%"]) {
      expect(isDateCode(code), code).toBe(false)
    }
  })

  it("built-in format ids, custom codes winning, locale date ids", () => {
    const formatFor = (id: number, custom: ReadonlyMap<number, string>) =>
      cellStyles([id], 1, new Map(Array.from(custom, ([k, code]) => [k, classify(code)]))).format(0)!
    expect(formatFor(0, new Map())).toEqual({ code: undefined, date: false })
    expect(formatFor(14, new Map())).toEqual({ code: "mm-dd-yy", date: true })
    expect(formatFor(4, new Map())).toEqual({ code: "#,##0.00", date: false })
    expect(formatFor(14, new Map([[14, "0.00"]]))).toEqual({ code: "0.00", date: false })
    expect(formatFor(31, new Map())).toEqual({ code: undefined, date: true })
    // the Thai built-in date/time ids
    for (let id = 71; id <= 81; id++) expect(formatFor(id, new Map()), String(id)).toEqual({ code: undefined, date: true })
    expect(formatFor(70, new Map()).date).toBe(false)
    expect(formatFor(82, new Map()).date).toBe(false)
    expect(formatFor(500, new Map())).toEqual({ code: undefined, date: false })
  })

  it("custom formats: a later code with the same id wins; the same code again keeps the classified format", () => {
    const formats = new Map<number, NumberFormat>()
    addFormat(formats, 164, "0.00")
    const first = formats.get(164)
    addFormat(formats, 164, "0.00")
    expect(formats.get(164)).toBe(first)
    addFormat(formats, 164, "yyyy")
    expect(formats.get(164)).toEqual({ code: "yyyy", date: true })
  })

  it("cell styles: a lookup by style index, undefined past the styles", () => {
    const styles = cellStyles(new Uint16Array([0, 14, 164, 0]), 3, new Map([[164, classify("0.000")]]))
    expect([styles.count, styles.format(0), styles.format(1)?.date, styles.format(2)]).toEqual([3, { code: undefined, date: false }, true, { code: "0.000", date: false }])
    // the id array may be longer than the styles kept
    expect(styles.format(3)).toBeUndefined()
    // a format is shared, not built per lookup
    expect(styles.format(2)).toBe(styles.format(2))
    expect(styles.format(1)).toBe(styles.format(1))
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

  it("_xHHHH_ escapes are decoded once, in <is>, in <v> of str and of inlineStr", async () => {
    const { sheets } = await readOk({
      sheets: [{
        name: "S",
        data: `<row><c t="inlineStr"><is><t>a_x0041_ _x005F_x0042_</t></is></c><c t="str"><v>b_x0043_</v></c><c t="inlineStr"><v>c_x0044_</v></c></row>`
      }]
    })
    expect(sheets[0]![0]!.cells.map(Cell.value)).toEqual(["aA _x0042_", "bC", "cD"])
  })

  it("_xHHHH_ escapes are decoded in sheet names and in defined names and their text, not in number format codes", async () => {
    const { wb, sheets } = await readOk({
      xfs: [0, 164],
      numFmts: { 164: "0.0_x0042_" },
      sheets: [{ name: "A_x0042_C", data: `<row><c s="1"><v>1</v></c></row>` }, { name: "_x005F_x0044_" }],
      definedNames: `<definedName name="N_x0042_" localSheetId="0">'A_x0042_C'!$A$1</definedName>`
    })
    expect(wb.sheets.map((s) => s.name)).toEqual(["ABC", "_x0044_"])
    expect(wb.definedNames).toEqual([{ name: "NB", formula: "'ABC'!$A$1", hidden: false, scope: 0 }])
    expect(sheets[0]![0]!.cells[0]!.numFmt).toBe("0.0_x0042_")
    // in a formula too
    const formulas = await readOk({ sheets: [{ name: "S", data: `<row><c><f>A_x0031_+_x005F_x0031_</f><v>1</v></c></row>` }] })
    expect(formulas.sheets[0]![0]!.cells[0]!.formula).toBe("A1+_x0031_")
    expect(values(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, "ABC"))))))
      .toEqual([[0, false, [[0, "Number", 1]]]])
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

  it("a custom number format given twice with the same id: the last one wins", async () => {
    const { sheets } = await readOk({
      xfs: [0, 164],
      numFmts: { 164: "0.00" },
      stylesExtra: `<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>`,
      sheets: [{ name: "S", data: `<row><c s="1"><v>45000</v></c></row>` }]
    })
    expect(sheets[0]![0]!.cells[0]).toEqual(Cell.date(0, new Date("2023-03-15T00:00:00.000Z"), { numFmt: "yyyy-mm-dd" }))
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
    // xsd:boolean: "true" too, but not "TRUE"
    const book = (value: string): TestWorkbook => ({
      xfs: [0, 14],
      workbookXml: `<workbookPr date1904="${value}"/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>`,
      sheets: [{ name: "S", data: `<row><c s="1"><v>0</v></c></row>` }]
    })
    expect(values((await readOk(book("true"))).sheets[0]!)[0]![2]).toEqual([[0, "Date", "1904-01-01T00:00:00.000Z"]])
    expect(values((await readOk(book("TRUE"))).sheets[0]!)[0]![2]).toEqual([[0, "Date", "1899-12-31T00:00:00.000Z"]])
  })

  it("sheet visibility, defined names, chart sheets, rows re-runnable", async () => {
    const bytes = await xlsx({
      sheets: [
        { name: "A", data: `<row><c><v>1</v></c></row>` },
        { name: "B", state: "hidden" },
        // parts that hold cell-like XML, but aren't worksheets: no rows, no details
        { name: "C", state: "veryHidden", kind: "chartsheet", data: `<row><c><v>9</v></c></row>` },
        { name: "D", kind: "dialogsheet", data: `<row><c><v>9</v></c></row>` },
        { name: "E", kind: "xlMacrosheet", data: `<row><c><v>9</v></c></row>` }
      ],
      definedNames: `<definedName name="Total">A!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="1" hidden="1">B!$A$1:$C$3</definedName><definedName name="Bad" localSheetId="7">#REF!</definedName>`
    })
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    expect(wb.sheets).toEqual([
      { name: "A", index: 0, visibility: "visible" },
      { name: "B", index: 1, visibility: "hidden" },
      { name: "C", index: 2, visibility: "veryHidden" },
      { name: "D", index: 3, visibility: "visible" },
      { name: "E", index: 4, visibility: "visible" }
    ])
    expect(wb.definedNames).toEqual([
      { name: "Total", formula: "A!$A$1", hidden: false },
      { name: "_xlnm.Print_Area", formula: "B!$A$1:$C$3", hidden: true, scope: 1 },
      { name: "Bad", formula: "#REF!", hidden: false }
    ])
    const rows = Workbook.rows(wb, "a")
    expect(values(Array.from(await Effect.runPromise(Stream.runCollect(rows))))).toEqual([[0, false, [[0, "Number", 1]]]])
    expect(Array.from(await Effect.runPromise(Stream.runCollect(rows))).length).toBe(1)
    for (const name of ["C", "D", "E"]) {
      expect(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, name)))), name).toEqual([])
      expect(await Effect.runPromise(Workbook.sheetDetails(wb, name)), name).toEqual({ hiddenColumns: [], merges: [] })
    }
  })

  it("a name scoped to the last sheet keeps its scope; one scoped at or past the sheet count is workbook-wide", async () => {
    const { wb } = await readOk({
      sheets: [{ name: "A" }, { name: "B" }, { name: "C", kind: "chartsheet" }],
      definedNames: [0, 2, 3, 4, 999999999].map((s) => `<definedName name="N${s}" localSheetId="${s}">A!$A$1</definedName>`).join("")
    })
    expect(wb.definedNames).toEqual([
      { name: "N0", formula: "A!$A$1", hidden: false, scope: 0 },
      { name: "N2", formula: "A!$A$1", hidden: false, scope: 2 },
      { name: "N3", formula: "A!$A$1", hidden: false },
      { name: "N4", formula: "A!$A$1", hidden: false },
      { name: "N999999999", formula: "A!$A$1", hidden: false }
    ])
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

  it("entry names with a leading slash read through Spreadsheet.read like through Xlsx.read", async () => {
    const bytes = await xlsx({ workbookPath: "/xl/workbook.xml", sheets: [{ name: "S", data: `<row><c><v>42</v></c></row>` }] })
    const cells = async (wb: Workbook.Workbook) => values(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0)))))
    const direct = await Effect.runPromise(Xlsx.read(bytes))
    expect(await cells(direct)).toEqual([[0, false, [[0, "Number", 42]]]])
    const wb = await Effect.runPromise(Spreadsheet.read(bytes))
    expect(wb.format).toBe("xlsx")
    expect(await cells(wb)).toEqual(await cells(direct))
  })

  it("limits from the layer at read time stay with the workbook", async () => {
    const bytes = await xlsx({ sheets: [{ name: "S", data: `<row><c><v>1</v></c></row><row><c><v>2</v></c></row><row><c><v>3</v></c></row>` }] })
    const wb = await Effect.runPromise(Xlsx.read(bytes).pipe(Effect.provide(SpreadsheetLimits.layer({ maxRows: 2 }))))
    const error = failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)))))
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 2 })
    expect(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0, { limits: { maxRows: 3 } })))).length).toBe(3)
  })
})

describe("xlsx: duplicate and missing sheet names", () => {
  const cell = (v: number) => `<row r="1"><c r="A1"><v>${v}</v></c></row>`
  const rowsOf = async (wb: Workbook.Workbook, sheet: string | number) =>
    values(Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, sheet)))))

  it("a sheet without a name gets the first Sheet<n>, from its position on, that no other sheet has in any case", async () => {
    const { wb } = await readOk({
      sheets: [{ name: "", data: cell(1) }, { name: "", data: cell(2) }, { name: "", data: cell(3) }, { name: "", data: cell(4) }],
      workbookXml: `<sheets><sheet sheetId="1" r:id="rId1"/><sheet name="Sheet1" sheetId="2" r:id="rId2"/>` +
        `<sheet name="SHEET2" sheetId="3" r:id="rId3"/><sheet sheetId="4" r:id="rId4"/></sheets>`
    })
    expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet3", "Sheet1", "SHEET2", "Sheet4"])
    expect(await rowsOf(wb, "Sheet1")).toEqual([[0, false, [[0, "Number", 2]]]])
    expect(await rowsOf(wb, "Sheet3")).toEqual([[0, false, [[0, "Number", 1]]]])
  })

  it("a later sheet with the same name gets a suffix; visibility, details and scoped names stay with their sheet", async () => {
    const hiddenColumn = sheetXml(cell(2)).replace("<sheetData>", `<cols><col min="2" max="2" hidden="1"/></cols><sheetData>`)
    const { wb } = await readOk({
      sheets: [{ name: "A", data: cell(1) }, { name: "A", raw: hiddenColumn, state: "hidden" }],
      definedNames: `<definedName name="N" localSheetId="1">A!$A$1</definedName>`
    })
    expect(wb.sheets).toEqual([{ name: "A", index: 0, visibility: "visible" }, { name: "A (2)", index: 1, visibility: "hidden" }])
    expect(wb.definedNames).toEqual([{ name: "N", formula: "A!$A$1", hidden: false, scope: 1 }])
    expect(await rowsOf(wb, "A")).toEqual([[0, false, [[0, "Number", 1]]]])
    expect(await rowsOf(wb, "A (2)")).toEqual([[0, false, [[0, "Number", 2]]]])
    expect(await rowsOf(wb, "a (2)")).toEqual([[0, false, [[0, "Number", 2]]]])
    expect((await Effect.runPromise(Workbook.sheetDetails(wb, "A (2)"))).hiddenColumns).toEqual([{ first: 1, last: 1 }])
    expect((await Effect.runPromise(Workbook.sheetDetails(wb, "A"))).hiddenColumns).toEqual([])
  })

  it("names are compared without case and after _xHHHH_ decoding, and a suffix skips names other sheets have", async () => {
    const names = ["Sheet1", "sheet1", "A", "_x0041_", "a (2)", "A"]
    const { wb } = await readOk({ sheets: names.map((name, k) => ({ name, data: cell(k) })) })
    expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1", "sheet1 (2)", "A", "A (3)", "a (2)", "A (4)"])
    for (const [k, s] of wb.sheets.entries()) expect(await rowsOf(wb, s.name), s.name).toEqual([[0, false, [[0, "Number", k]]]])
  })
})

describe("xlsx: row order", () => {
  const sheet = (data: string): TestWorkbook => ({ sheets: [{ name: "S", data }] })
  const malformedAt = (reason: string) => ({ _tag: "MalformedFile", format: "xlsx", reason })

  it("a row index given twice is MalformedFile, not two rows with the same index", async () => {
    expect(await errorOf(sheet(`<row r="1"><c r="A1"><v>1</v></c></row><row r="1"><c r="B1"><v>2</v></c></row>`)))
      .toMatchObject(malformedAt("row 1 after row 1"))
    // the second one hidden: not a second row 1 with its own hidden flag
    expect(await errorOf(sheet(`<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="1" hidden="1"><c r="B1"><v>3</v></c></row>`)))
      .toMatchObject(malformedAt("row 1 after row 2"))
  })

  it("a row index going back is MalformedFile, also when a row without r takes it from its first cell", async () => {
    expect(await errorOf(sheet(`<row r="3"><c r="A3"><v>1</v></c></row><row r="1"><c r="A1"><v>2</v></c></row>`)))
      .toMatchObject(malformedAt("row 1 after row 3"))
    expect(await errorOf(sheet(`<row r="5"><c r="A5"><v>1</v></c></row><row><c r="A2"><v>2</v></c></row>`)))
      .toMatchObject(malformedAt("row 2 after row 5"))
  })

  it("toRecords with header: 0 fails on rows out of order instead of reading the header row as data", async () => {
    const bytes = await xlsx(sheet(`<row r="2"><c r="A2"><v>1</v></c></row><row r="1"><c r="A1" t="inlineStr"><is><t>name</t></is></c></row>`))
    const exit = await Effect.runPromise(Effect.exit(Effect.gen(function*() {
      const wb = yield* Xlsx.read(bytes)
      return Array.from(yield* Stream.runCollect(Spreadsheet.toRecords(Workbook.rows(wb, 0), { header: 0 })))
    })))
    expect(failureOf(exit)).toMatchObject(malformedAt("row 1 after row 2"))
  })

  it("a row without r whose first cell has no r is the previous row + 1, not a later cell's row", async () => {
    // the later cell names row 5, so it isn't in this row (row 1)
    expect(await errorOf(sheet(`<row><c><v>1</v></c><c r="B5"><v>2</v></c></row>`))).toMatchObject(malformedAt(`cell "B5" in row 1`))
    expect(await errorOf(sheet(`<row r="3"/><row><c><v>1</v></c><c r="B7"><v>2</v></c></row>`))).toMatchObject(malformedAt(`cell "B7" in row 4`))
  })

  it("a row without r starts afresh: the previous row + 1, columns from A, not hidden", async () => {
    const { sheets } = await readOk(sheet(
      `<row r="2" hidden="1"><c r="C2"><v>1</v></c></row><row><c><v>2</v></c><c><v>3</v></c></row>` +
        `<row r="6" hidden="1"><c r="D6"><v>4</v></c></row><row/><row><c r="A9"><v>5</v></c></row>`
    ))
    expect(values(sheets[0]!)).toEqual([
      [1, true, [[2, "Number", 1]]],
      [2, false, [[0, "Number", 2], [1, "Number", 3]]],
      [5, true, [[3, "Number", 4]]],
      [6, false, []],
      [8, false, [[0, "Number", 5]]]
    ])
  })

  it("a cell whose reference names another row than its row is MalformedFile", async () => {
    expect(await errorOf(sheet(`<row r="1"><c r="A5"><v>1</v></c></row>`))).toMatchObject(malformedAt(`cell "A5" in row 1`))
    // a row without r takes its first cell's row; a later cell must agree
    expect(await errorOf(sheet(`<row><c r="A2"><v>1</v></c><c r="B7"><v>2</v></c></row>`))).toMatchObject(malformedAt(`cell "B7" in row 2`))
    // references without a row (column only) are fine
    const { sheets } = await readOk(sheet(`<row r="2"><c r="A"><v>1</v></c><c r="B2"><v>2</v></c></row><row r="4"><c r="C4"><v>3</v></c></row>`))
    expect(values(sheets[0]!)).toEqual([[1, false, [[0, "Number", 1], [1, "Number", 2]]], [3, false, [[2, "Number", 3]]]])
  })

  it("a bad row past the rows taken doesn't fail an early stop", async () => {
    // enough rows before the bad one that it comes in a later inflated piece
    let data = ""
    for (let r = 1; r <= 5000; r++) data += `<row r="${r}"><c r="A${r}"><v>${r}</v></c></row>`
    data += `<row r="1"><c r="B1"><v>0</v></c></row>`
    const bytes = await xlsx(sheet(data))
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)))))).toMatchObject(malformedAt("row 1 after row 5000"))
    const first = Array.from(await Effect.runPromise(Stream.runCollect(Stream.take(Workbook.rows(wb, 0), 2))))
    expect(values(first)).toEqual([[0, false, [[0, "Number", 1]]], [1, false, [[0, "Number", 2]]]])
  })
})

describe("xlsx: an element nested in one of its own kind", () => {
  const sheet = (data: string): TestWorkbook => ({ sheets: [{ name: "S", data }] })
  const malformedIn = (reason: string, part: string) => ({ _tag: "MalformedFile", format: "xlsx", reason, part })
  /** A package with two worksheets whose workbook part holds `body`. */
  const withWorkbook = async (body: string) => readExit(await xlsx({ sheets: [{ name: "A" }, { name: "B" }], workbookXml: body }))

  it("a <row> inside a row is MalformedFile, not a row replacing the open one", async () => {
    expect(await errorOf(sheet(`<row r="1"><c r="A1"><v>1</v></c><row r="2"><c r="A2"><v>2</v></c></row></row><row r="3"><c r="A3"><v>3</v></c></row>`)))
      .toMatchObject(malformedIn("<row> inside a <row>", "xl/worksheets/sheet1.xml"))
    // also below another element of the row, and in a row without r
    expect(await errorOf(sheet(`<row><c><v>1</v></c><x><row/></x></row>`))).toMatchObject(malformedIn("<row> inside a <row>", "xl/worksheets/sheet1.xml"))
    // the rows before it still come out first
    const bytes = await xlsx(sheet(`<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><row r="3"/></row>`))
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    const seen: Array<number> = []
    const exit = await Effect.runPromiseExit(Stream.runForEach(Workbook.rows(wb, 0), (row) => Effect.sync(() => seen.push(row.index))))
    expect(seen).toEqual([0])
    expect(failureOf(exit)).toMatchObject({ _tag: "MalformedFile" })
  })

  it("a <sheetData> inside the sheet data is MalformedFile, not the end of the sheet data", async () => {
    expect(await errorOf(sheet(`<sheetData></sheetData><row r="1"><c r="A1"><v>1</v></c></row>`)))
      .toMatchObject(malformedIn("<sheetData> inside a <sheetData>", "xl/worksheets/sheet1.xml"))
    expect(await errorOf(sheet(`<row r="1"><sheetData/></row>`))).toMatchObject(malformedIn("<sheetData> inside a <sheetData>", "xl/worksheets/sheet1.xml"))
  })

  it("an <si> inside a shared string is MalformedFile, not a string replacing the open one", async () => {
    const error = await errorOf({
      sharedStringsXml: `<sst xmlns="x"><si><t>a</t><si><t>b</t></si></si><si><t>c</t></si></sst>`,
      sheets: [{ name: "S", data: `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>` }]
    })
    expect(error).toMatchObject(malformedIn("<si> inside an <si>", "xl/sharedStrings.xml"))
  })

  it("a <sheets> inside <sheets> is MalformedFile, not the end of the sheet list", async () => {
    const exit = await withWorkbook(`<sheets><sheets></sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets>`)
    expect(failureOf(exit)).toMatchObject(malformedIn("<sheets> inside <sheets>", "xl/workbook.xml"))
    // after a sheet: not a workbook with only the sheets before it
    const after = await withWorkbook(`<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheets/><sheet name="B" sheetId="2" r:id="rId2"/></sheets>`)
    expect(failureOf(after)).toMatchObject(malformedIn("<sheets> inside <sheets>", "xl/workbook.xml"))
    // the same list without the nested element reads both sheets
    const ok = await withWorkbook(`<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets>`)
    expect(Exit.isSuccess(ok) && ok.value.wb.sheets.map((s) => s.name)).toEqual(["A", "B"])
  })

  it("a <definedName> inside a defined name is MalformedFile, not a name replacing the open one", async () => {
    expect(await errorOf({ sheets: [{ name: "S" }], definedNames: `<definedName name="A">x<definedName name="B">y</definedName>z</definedName>` }))
      .toMatchObject(malformedIn("<definedName> inside a <definedName>", "xl/workbook.xml"))
  })

  it("a phonetic run inside a <t> is skipped, and the <t>'s text after it is kept", async () => {
    // the <t> inside the phonetic run neither adds its text nor ends the outer <t>
    const { sheets } = await readOk({
      sharedStringsXml: `<sst xmlns="x"><si><t>a<rPh sb="0" eb="1"><t>b</t></rPh>c</t></si><si><t>d</t></si></sst>`,
      sheets: [{ name: "S", data: `<row><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c t="inlineStr"><is><t>e<rPh><t>f</t></rPh>g</t></is></c></row>` }]
    })
    expect(sheets[0]![0]!.cells.map(Cell.value)).toEqual(["ac", "d", "eg"])
  })
})

describe("xlsx: where the workbook, shared-strings and styles elements are read", () => {
  const cells = `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" s="0"><v>45000</v></c><c r="D1" s="1"><v>45000</v></c><c r="E1" s="2"><v>45000</v></c></row>`
  it("under any wrapper: sheets, defined names, shared strings, number formats and cell styles", async () => {
    const wrap = (s: string) => `<mc:AlternateContent xmlns:mc="m"><mc:Choice>${s}</mc:Choice></mc:AlternateContent>`
    const { wb, sheets } = await readOk({
      sheets: [{ name: "A", data: cells }, { name: "B" }],
      workbookXml: `${wrap(`<workbookPr date1904="1"/>`)}<sheets>${wrap(`<sheet name="A" r:id="rId1"/>`)}<sheet name="B" r:id="rId2"/></sheets>` +
        wrap(`<definedNames>${wrap(`<definedName name="N">S!A1</definedName>`)}</definedNames>`),
      sharedStringsXml: `<sst xmlns="x">${wrap(`<si><r><t>o</t></r>${wrap(`<t>ne</t>`)}</si>`)}<si><t>two</t></si></sst>`,
      xfs: [],
      stylesExtra: `<numFmts>${wrap(`<numFmt numFmtId="200" formatCode="0.0"/>`)}</numFmts><cellXfs>${wrap(`<xf numFmtId="200"/>`)}<xf numFmtId="14"/></cellXfs>`
    })
    expect(wb.sheets.map((s) => s.name)).toEqual(["A", "B"])
    expect(wb.definedNames).toEqual([expect.objectContaining({ name: "N", formula: "S!A1" })])
    const row = sheets[0]![0]!.cells
    expect(row.slice(0, 2).map(Cell.value)).toEqual(["one", "two"])
    // xf 0 has the custom format; xf 1 a date in the 1904 system; xf 2 is past the styles
    expect(row.slice(2).map((c) => [c._tag, c.numFmt])).toEqual([["Number", "0.0"], ["Date", "mm-dd-yy"], ["Number", undefined]])
    expect((Cell.value(row[3]!) as Date).toISOString()).toBe("2027-03-16T00:00:00.000Z")
  })

  it("only inside their lists: a <sheet>, <numFmt> or <xf> elsewhere is ignored", async () => {
    const { wb, sheets } = await readOk({
      sheets: [{ name: "A", data: cells }, { name: "B" }],
      workbookXml: `<sheet name="X" r:id="rId2"/><sheets><sheet name="A" r:id="rId1"/><sheet name="B" r:id="rId2"/></sheets><x><sheet name="Y" r:id="rId2"/></x>`,
      sharedStrings: ["x", "y"],
      xfs: [],
      stylesExtra: `<numFmt numFmtId="200" formatCode="0.0"/><dxfs><dxf><numFmt numFmtId="201" formatCode="0.00"/></dxf></dxfs><xf numFmtId="14"/>` +
        `<cellStyleXfs><xf numFmtId="14"/></cellStyleXfs><cellXfs><xf numFmtId="200"/><xf numFmtId="201"/></cellXfs>`
    })
    expect(wb.sheets.map((s) => s.name)).toEqual(["A", "B"])
    // 200 and 201 are not number formats of the workbook: no format
    expect(sheets[0]![0]!.cells.slice(2).map((c) => [c._tag, c.numFmt])).toEqual([["Number", undefined], ["Number", undefined], ["Number", undefined]])
  })

  it("a list nested in a list of its own kind: the entries after it in the outer list are read too", async () => {
    const { sheets } = await readOk({
      sheets: [{ name: "S", data: cells }],
      sharedStrings: ["x", "y"],
      xfs: [],
      stylesExtra: `<numFmts><numFmts><numFmt numFmtId="200" formatCode="0.0"/></numFmts><numFmt numFmtId="201" formatCode="yyyy-mm-dd"/></numFmts>` +
        `<cellXfs><cellXfs><xf numFmtId="200"/></cellXfs><xf numFmtId="201"/><xf numFmtId="14"/></cellXfs>`
    })
    expect(sheets[0]![0]!.cells.slice(2).map((c) => [c._tag, c.numFmt])).toEqual([["Number", "0.0"], ["Date", "yyyy-mm-dd"], ["Date", "mm-dd-yy"]])
  })
})

describe("xlsx: rows completed before an error", () => {
  const sheet = (data: string): TestWorkbook => ({ sheets: [{ name: "S", data }] })
  /** The rows a consumer sees before the stream ends or fails, and the error. */
  const untilError = async (bytes: Uint8Array, limits?: LimitOverrides) => {
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    const seen: Array<Workbook.Row> = []
    const exit = await Effect.runPromiseExit(Stream.runForEach(Workbook.rows(wb, 0, { limits }), (row) => Effect.sync(() => seen.push(row))))
    return { indexes: seen.map((row) => row.index), error: failureOf(exit) }
  }
  const row = (r: number, cells = `<c r="A${r}"><v>${r}</v></c>`) => `<row r="${r}">${cells}</row>`
  const tenRows = Array.from({ length: 10 }, (_, k) => row(k + 1)).join("")

  it("come out before the error when they share its piece", async () => {
    const bad = await untilError(await xlsx(sheet(tenRows + row(4))))
    expect(bad.indexes).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9])
    expect(bad.error).toMatchObject({ _tag: "MalformedFile", reason: "row 4 after row 10" })
    const limited = await untilError(await xlsx(sheet(tenRows)), { maxRows: 5 })
    expect(limited.indexes).toEqual([0, 1, 2, 3, 4])
    expect(limited.error).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows" })
  })

  it("don't depend on where the part is split into pieces", async () => {
    // the same rows and bad row, in one piece, or with a row of 10 000
    // cells (about 200 KB of XML) that puts rows 1-5 in an earlier piece
    let filler = ""
    for (let k = 0; k < 10_000; k++) filler += `<c><v>${k}</v></c>`
    const small = await untilError(await xlsx(sheet(tenRows + row(4))))
    const big = await untilError(await xlsx(sheet(Array.from({ length: 10 }, (_, k) => k === 5 ? row(6, filler) : row(k + 1)).join("") + row(4))))
    expect(small.indexes).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9])
    expect(small.error).toMatchObject({ _tag: "MalformedFile", reason: "row 4 after row 10" })
    expect(big).toEqual(small)
  })

  it("can be taken without the error when it is in their piece", async () => {
    const wb = await Effect.runPromise(Xlsx.read(await xlsx(sheet(tenRows + row(4)))))
    const first = Array.from(await Effect.runPromise(Stream.runCollect(Stream.take(Workbook.rows(wb, 0), 3))))
    expect(values(first)).toEqual([[0, false, [[0, "Number", 1]]], [1, false, [[0, "Number", 2]]], [2, false, [[0, "Number", 3]]]])
  })
})

describe("xlsx: where a cell's value comes from", () => {
  const firstCells = async (data: string) => {
    const { sheets } = await readOk({ sharedStrings: ["sst0", "sst1"], sheets: [{ name: "S", data: `<row r="1">${data}</row>` }] })
    return sheets[0]![0]!.cells.map((c) => [c.column, c._tag, Cell.value(c), c.formula])
  }

  it("the cell's type decides between <v> and <is>", async () => {
    // an <is> doesn't turn a shared string, formula string, number or boolean into inline text
    expect(await firstCells(`<c r="A1" t="s"><v>0</v><is><t>x</t></is></c>`)).toEqual([[0, "Text", "sst0", undefined]])
    expect(await firstCells(`<c r="A1" t="s"><is><t>x</t></is><v>1</v></c>`)).toEqual([[0, "Text", "sst1", undefined]])
    expect(await firstCells(`<c r="A1" t="str"><f>"a"</f><v>b</v><is><t>x</t></is></c>`)).toEqual([[0, "Text", "b", `"a"`]])
    expect(await firstCells(`<c r="A1"><v>3</v><is><t>x</t></is></c>`)).toEqual([[0, "Number", 3, undefined]])
    expect(await firstCells(`<c r="A1"><is><t>x</t></is><v>3</v></c>`)).toEqual([[0, "Number", 3, undefined]])
    expect(await firstCells(`<c r="A1" t="b"><v>1</v><is><t>x</t></is></c>`)).toEqual([[0, "Boolean", true, undefined]])
    // an inline string takes <is>, before or after a <v>, which isn't glued onto it
    expect(await firstCells(`<c r="A1" t="inlineStr"><is><t>x</t></is><v>5</v></c>`)).toEqual([[0, "Text", "x", undefined]])
    expect(await firstCells(`<c r="A1" t="inlineStr"><v>5</v><is><t>x</t></is></c>`)).toEqual([[0, "Text", "x", undefined]])
    expect(await firstCells(`<c r="A1" t="inlineStr"><is><t>a</t></is><v>_x0041_</v></c>`)).toEqual([[0, "Text", "a", undefined]])
    // still read: an inline string written in <v>, and an <is> in a cell without a type
    expect(await firstCells(`<c r="A1" t="inlineStr"><v>c_x0044_</v></c>`)).toEqual([[0, "Text", "cD", undefined]])
    expect(await firstCells(`<c r="A1"><is><t>7</t></is></c>`)).toEqual([[0, "Text", "7", undefined]])
  })

  it("a repeated <v>, <f> or <is> in one cell: the first one counts, nothing is glued together", async () => {
    expect(await firstCells(`<c r="A1"><v>1</v><v>2</v></c>`)).toEqual([[0, "Number", 1, undefined]])
    expect(await firstCells(`<c r="A1"><v></v><v>2</v></c>`)).toEqual([[0, "Empty", undefined, undefined]])
    expect(await firstCells(`<c r="A1" t="str"><v>a</v><v>b</v></c>`)).toEqual([[0, "Text", "a", undefined]])
    expect(await firstCells(`<c r="A1" t="s"><v>0</v><v>1</v></c>`)).toEqual([[0, "Text", "sst0", undefined]])
    expect(await firstCells(`<c r="A1"><f>1+1</f><f>2</f><v>2</v></c>`)).toEqual([[0, "Number", 2, "1+1"]])
    expect(await firstCells(`<c r="A1" t="inlineStr"><is><t>x</t></is><is><t>y</t></is></c>`)).toEqual([[0, "Text", "x", undefined]])
  })

  it("elements nested in a cell (as in its extLst) neither end the cell nor give its value", async () => {
    expect(await firstCells(`<c r="A1"><extLst><ext uri="x"><c/></ext></extLst><v>1</v></c>`)).toEqual([[0, "Number", 1, undefined]])
    expect(await firstCells(`<c r="A1"><extLst><ext uri="x"><c r="Z9"><v>7</v></c></ext></extLst><v>1</v></c><c r="B1"><v>2</v></c>`))
      .toEqual([[0, "Number", 1, undefined], [1, "Number", 2, undefined]])
    expect(await firstCells(`<c r="A1"><v>1</v><extLst><ext uri="x"><v>9</v><f>Z</f></ext></extLst></c>`)).toEqual([[0, "Number", 1, undefined]])
    expect(await firstCells(`<c r="A1" t="inlineStr"><extLst><ext uri="x"><is><t>q</t></is></ext></extLst><is><t>x</t></is></c>`))
      .toEqual([[0, "Text", "x", undefined]])
    // text around an element nested in the <v> or <f> is kept, the element's own text isn't
    expect(await firstCells(`<c r="A1"><f>1<x>9</x>+1</f><v>1<x>9</x>2</v></c>`)).toEqual([[0, "Number", 12, "1+1"]])
    // an inline string's text spans its runs
    expect(await firstCells(`<c r="A1" t="inlineStr"><is><r><t>a</t></r><r><rPr/><t>b</t></r></is><v>9</v></c>`)).toEqual([[0, "Text", "ab", undefined]])
    // an element named is inside the inline string doesn't end it
    expect(await firstCells(`<c r="A1" t="inlineStr"><is><t>x</t><extLst><is/></extLst><t>y</t></is></c>`)).toEqual([[0, "Text", "xy", undefined]])
  })
})

describe("xlsx: integer attributes, each with its own rule for a missing or bad value", () => {
  it("a defined name's localSheetId: up to nine digits; a bad one is workbook-wide", async () => {
    const { wb } = await readOk({
      sheets: [{ name: "A" }, { name: "B" }],
      definedNames: ["000000001", "0000000001", "x", "-1", " 1", ""].map((s, k) => `<definedName name="N${k}" localSheetId="${s}">A!$A$1</definedName>`).join("")
    })
    expect(wb.definedNames.map((n) => n.scope)).toEqual([1, undefined, undefined, undefined, undefined, undefined])
  })

  it("numFmtId: up to nine digits; a bad one skips a custom code, and is General in a cell style", async () => {
    const { sheets } = await readOk({
      // custom 164 written with leading zeros; a bad id's code is not added under any id
      numFmts: { "000000164": "yyyy-mm-dd", "x": "0.00", "1000000000": "0.0" },
      xfs: ["000000164", "x", "0000000014", "-1", "1"],
      sheets: [{ name: "S", data: `<row r="1">${[0, 1, 2, 3, 4].map((s) => `<c s="${s}"><v>45000</v></c>`).join("")}</row>` }]
    })
    expect(sheets[0]![0]!.cells.map((c) => [c._tag, c.numFmt])).toEqual([
      ["Date", "yyyy-mm-dd"],
      ["Number", undefined],
      ["Number", undefined],
      ["Number", undefined],
      ["Number", "0"]
    ])
  })

  it("a <numFmt> without a code is skipped: it doesn't replace an earlier one with its id", async () => {
    const { sheets } = await readOk({
      numFmts: { 164: "yyyy-mm-dd" },
      // a second <numFmts>: its entries count like the first one's
      stylesExtra: `<numFmts><numFmt numFmtId="164"/></numFmts>`,
      xfs: [164],
      sheets: [{ name: "S", data: `<row r="1"><c s="0"><v>45000</v></c></row>` }]
    })
    expect(sheets[0]![0]!.cells.map((c) => [c._tag, c.numFmt])).toEqual([["Date", "yyyy-mm-dd"]])
  })

  it("a sheet's relationship id: any prefix counts, whatever namespace it is bound to", async () => {
    const { wb, sheets } = await readOk({
      workbookXml: `<sheets><sheet name="A" sheetId="1" xmlns:x="urn:not-relationships" x:id="rId2"/><sheet name="B" sheetId="2" r:id="rId1"/></sheets>`,
      sheets: [{ name: "S1", data: `<row r="1"><c><v>1</v></c></row>` }, { name: "S2", data: `<row r="1"><c><v>2</v></c></row>` }]
    })
    expect(wb.sheets.map((s) => s.name)).toEqual(["A", "B"])
    expect(sheets.map((rows) => values(rows))).toEqual([[[0, false, [[0, "Number", 2]]]], [[0, false, [[0, "Number", 1]]]]])
  })

  it("a cell's style index: up to nine digits; a missing or bad one is no style", async () => {
    const { sheets } = await readOk({
      xfs: [14, 14],
      sheets: [{ name: "S", data: `<row r="1">${["000000001", "x", "-1", "", "0000000001"].map((s) => `<c s="${s}"><v>45000</v></c>`).join("")}<c><v>45000</v></c></row>` }]
    })
    expect(sheets[0]![0]!.cells.map((c) => c._tag)).toEqual(["Date", "Number", "Number", "Number", "Number", "Number"])
  })

  it("a row number: up to ten digits; more digits are past the row limit, anything else is MalformedFile", async () => {
    const sheet = (data: string): TestWorkbook => ({ sheets: [{ name: "S", data }] })
    expect(values((await readOk(sheet(`<row r="0000000001"><c><v>1</v></c></row>`))).sheets[0]!)).toEqual([[0, false, [[0, "Number", 1]]]])
    expect(await errorOf(sheet(`<row r="9999999999"/>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", actual: 9999999999 })
    expect(await errorOf(sheet(`<row r="00000000001"/>`))).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", actual: Infinity })
    for (const r of ["", "x", "-1", "+1", " 1", "1.0", "\u0663"]) {
      expect(await errorOf(sheet(`<row r="${r}"/>`)), r).toMatchObject({ _tag: "MalformedFile", reason: `bad row number ${JSON.stringify(r)}` })
    }
    expect(await errorOf(sheet(`<row r="000"/>`))).toMatchObject({ _tag: "MalformedFile", reason: "row number 0" })
  })
})

describe("xlsx: sheet details", () => {
  const withLayout = (layout: string, data = `<row><c><v>1</v></c></row>`) =>
    xlsx({ sheets: [{ name: "S", raw: `<worksheet xmlns="x"><dimension ref="B2:D9"/><cols>${layout}</cols><sheetData>${data}</sheetData><mergeCells><mergeCell ref="A1:C1"/><mergeCell ref="B3:B4"/></mergeCells></worksheet>` }] })
  const details = async (bytes: Uint8Array, limits?: LimitOverrides) =>
    Effect.runPromise(Effect.exit(Effect.flatMap(Xlsx.read(bytes, { limits }), (wb) => Workbook.sheetDetails(wb, 0))))

  it("hidden columns as merged spans, merges, the declared dimension", async () => {
    // out of order on purpose: spans are sorted before they are merged
    const exit = await details(await withLayout(`<col min="12" max="12" hidden="1"/><col min="2" max="3" hidden="1"/><col min="4" max="4" hidden="true"/><col min="6" max="6" width="9"/><col min="8" max="20000" hidden="1"/><col min="9" max="9" hidden="1"/>`))
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
    // both ends are needed, with up to nine digits each
    for (const range of [`min="1"`, `max="1"`, `min="1" max="1000000000"`, `min="0000000001" max="1"`, `min="1" max=""`]) {
      expect(failureOf(await details(await withLayout(`<col ${range} hidden="1"/>`))), range).toMatchObject({ _tag: "MalformedFile" })
    }
    const nineDigits = await details(await withLayout(`<col min="000000002" max="000000003" hidden="1"/>`))
    expect(Exit.isSuccess(nineDigits) && nineDigits.value.hiddenColumns).toEqual([{ first: 1, last: 2 }])
    // a span that isn't hidden isn't read
    for (const hidden of ["", ` hidden="0"`, ` hidden="false"`, ` hidden="TRUE"`]) {
      const exit = await details(await withLayout(`<col min="x"${hidden}/>`))
      expect(Exit.isSuccess(exit) && exit.value.hiddenColumns, hidden).toEqual([])
    }
    const badMerge = await xlsx({ sheets: [{ name: "S", raw: `<worksheet xmlns="x"><sheetData/><mergeCells><mergeCell ref="A1:"/></mergeCells></worksheet>` }] })
    expect(failureOf(await details(badMerge))).toMatchObject({ _tag: "MalformedFile" })
    expect(failureOf(await details(await withLayout(`<col min="1" max="1"/>`.repeat(5)), { maxColumns: 4 }))).toMatchObject({ limit: "maxColumns" })
    expect(failureOf(await details(await withLayout(""), { maxCellsPerSheet: 1 }))).toMatchObject({ limit: "maxCellsPerSheet" })
  })

  it("a <sheetData> inside the sheet data is MalformedFile, as for rows, not the end of the sheet data", async () => {
    const malformed = { _tag: "MalformedFile", format: "xlsx", reason: "<sheetData> inside a <sheetData>", part: "xl/worksheets/sheet1.xml" }
    // layout after the inner end tag is still inside the outer sheet data
    const nested = await xlsx({
      sheets: [{
        name: "S",
        raw: `<worksheet xmlns="x"><sheetData><sheetData/><cols><col min="1" max="1" hidden="1"/></cols><mergeCells><mergeCell ref="A1:B2"/></mergeCells></sheetData></worksheet>`
      }]
    })
    expect(failureOf(await details(nested))).toMatchObject(malformed)
    const wb = await Effect.runPromise(Xlsx.read(nested))
    expect(failureOf(await Effect.runPromiseExit(Stream.runCollect(Workbook.rows(wb, 0))))).toMatchObject(malformed)
    // also inside a row
    expect(failureOf(await details(await withLayout("", `<row r="1"><sheetData/></row>`)))).toMatchObject(malformed)
  })

  it("a cell's content is skipped, as for rows, whatever its element names", async () => {
    // a <sheetData> (in any namespace) and a <mergeCell> in a cell's extension are neither nesting nor layout
    const bytes = await withLayout(
      `<col min="2" max="2" hidden="1"/>`,
      `<row r="1"><c r="A1"><v>1</v><extLst><ext uri="x" xmlns:y="y"><y:sheetData/><mergeCell ref="E5:F6"/><col min="1" max="1" hidden="1"/></ext></extLst></c></row>`
    )
    const exit = await details(bytes)
    expect(Exit.isSuccess(exit) && exit.value).toEqual({
      hiddenColumns: [{ first: 1, last: 1 }],
      merges: [
        { firstRow: 0, firstColumn: 0, lastRow: 0, lastColumn: 2 },
        { firstRow: 2, firstColumn: 1, lastRow: 3, lastColumn: 1 }
      ],
      dimension: { firstRow: 1, firstColumn: 1, lastRow: 8, lastColumn: 3 }
    })
    const read = await readExit(bytes)
    expect(Exit.isSuccess(read) && values(read.value.sheets[0]!)).toEqual([[0, false, [[0, "Number", 1]]]])
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
    // rows repeating an index don't pile up: the first repeat fails
    expect(await errorOf(sheet(`<row r="1"/>`.repeat(5)), { maxRows: 3 })).toMatchObject({ _tag: "MalformedFile", reason: "row 1 after row 1" })
  })

  it("the input size and defined names have limits", async () => {
    const bytes = await xlsx({ sheets: [{ name: "S" }] })
    const error = failureOf(await readExit(bytes, { maxInputBytes: 10 }))
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxInputBytes", max: 10, actual: bytes.length })
    // text runs split by comments: each is short, together they're too long
    const longName = await errorOf({ sheets: [{ name: "S" }], definedNames: `<definedName name="N">${"A1,<!-- -->".repeat(100)}</definedName>` }, { maxStringLength: 100 })
    expect(longName).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", max: 100, actual: 102 })
  })

  it("sheets, shared strings and string lengths have limits", async () => {
    expect(await errorOf({ sheets: [{ name: "A" }, { name: "B" }, { name: "C" }] }, { maxSheets: 2 })).toMatchObject({ limit: "maxSheets", max: 2 })
    expect(await errorOf({ sharedStrings: ["a", "b", "c"], sheets: [{ name: "S" }] }, { maxSharedStrings: 2 })).toMatchObject({ limit: "maxSharedStrings" })
    expect(await errorOf({ sharedStringsXml: `<sst xmlns="x"><si>${"<r><t>abcd</t></r>".repeat(50)}</si></sst>`, sheets: [{ name: "S" }] }, { maxStringLength: 100 })).toMatchObject({ limit: "maxStringLength" })
    expect(await errorOf(sheet(`<row><c><f>${"x".repeat(500)}</f></c></row>`), { maxStringLength: 100 })).toMatchObject({ limit: "maxStringLength" })
  })

  it("a string's length counts all its text runs, and only the runs that are part of it", async () => {
    const max = { maxStringLength: 100 }
    // the error comes at the run that goes past the limit, with the total so far
    const atRun = { _tag: "LimitExceeded", limit: "maxStringLength", max: 100 }
    expect(await errorOf({ sharedStringsXml: `<sst xmlns="x"><si>${"<r><t>abcd</t></r>".repeat(50)}</si></sst>`, sheets: [{ name: "S" }] }, max))
      .toMatchObject({ ...atRun, actual: 104 })
    expect(await errorOf(sheet(`<row><c t="inlineStr"><is>${"<t>abcd</t>".repeat(30)}</is></c></row>`), max)).toMatchObject({ ...atRun, actual: 104 })
    expect(await errorOf(sheet(`<row><c><v>${"1234567<!-- -->".repeat(20)}</v></c></row>`), max)).toMatchObject({ ...atRun, actual: 105 })
    expect(await errorOf(sheet(`<row><c><f>${"1+2+3<x/>".repeat(30)}</f></c></row>`), max)).toMatchObject({ ...atRun, actual: 105 })
    // phonetic runs, text outside <t> and elements nested in a <v> or <f> don't count
    const long = (c: string) => c.repeat(60)
    const { sheets } = await readOk({
      // each item counts its own text
      sharedStringsXml: `<sst xmlns="x"><si>${long("z")}<t>${long("a")}<rPh><t>${long("b")}</t></rPh>${"c".repeat(30)}</t></si><si><t>${long("d")}</t></si></sst>`,
      sheets: [{ name: "S", data: `<row><c t="s"><v>0</v></c><c t="str"><f>${long("1")}<x>${long("9")}</x></f><v>${long("2")}<x>${long("8")}</x></v></c></row>` }]
    }, max)
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Text", long("a") + "c".repeat(30)], [1, "Text", long("2")]]]])
    // a phonetic run's text is left out, but an <si> inside it is still one inside the item
    expect(await errorOf({ sharedStringsXml: `<sst xmlns="x"><si><t>a</t><rPh><si/></rPh></si></sst>`, sheets: [{ name: "S" }] }))
      .toMatchObject({ _tag: "MalformedFile", reason: "<si> inside an <si>" })
    expect(sheets[0]![0]!.cells[1]!.formula).toBe(long("1"))
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
        reason: "sheet part xl/worksheets/sheet2.xml is missing",
        part: "xl/workbook.xml"
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

  describe("sheets sharing one worksheet part", () => {
    const pkg = (sheets: string, rels: string, sheetData = `<row><c><v>1</v></c></row>`) =>
      buildZip([
        { name: "_rels/.rels", data: enc(`<Relationships xmlns="x"><Relationship Id="r" Type="http://x/officeDocument" Target="xl/workbook.xml"/></Relationships>`) },
        { name: "xl/workbook.xml", data: enc(`<workbook xmlns="x" xmlns:r="y"><sheets>${sheets}</sheets></workbook>`) },
        { name: "xl/_rels/workbook.xml.rels", data: enc(`<Relationships xmlns="x">${rels}</Relationships>`) },
        { name: "xl/worksheets/sheet1.xml", data: enc(sheetXml(sheetData)) }
      ])
    const rel = (id: string, target = "worksheets/sheet1.xml") => `<Relationship Id="${id}" Type="http://x/worksheet" Target="${target}"/>`
    const onOneId = (count: number) => Array.from({ length: count }, (_, i) => `<sheet name="S${i}" r:id="rId1"/>`).join("")
    const rowsExit = (wb: Workbook.Workbook, sheet: string | number, limits?: LimitOverrides) =>
      Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, sheet, limits === undefined ? undefined : { limits }))))

    it("open, and each sheet reads the part", async () => {
      // N sheets on one r:id
      const sameId = await readExit(await pkg(onOneId(50), rel("rId1")))
      expect(Exit.isSuccess(sameId) && sameId.value.sheets.map((rows) => values(rows))).toEqual(Array(50).fill([[0, false, [[0, "Number", 1]]]]))
      // two relationships with the same target (written differently)
      const sameTarget = await readExit(await pkg(`<sheet name="A" r:id="rId1"/><sheet name="B" r:id="rId2"/>`, rel("rId1") + rel("rId2", "/XL/worksheets/sheet1.xml")))
      expect(Exit.isSuccess(sameTarget) && sameTarget.value.sheets.map((rows) => rows.length)).toEqual([1, 1])
    })

    it("each pay for what they read in the inflate budget, so aliasing a part amplifies nothing", async () => {
      // a sheet part of about 50 KB; a budget for the open-time reads and three and a half of it
      const sheetData = `<row><c><v>1</v></c></row>${" ".repeat(50_000)}`
      const wb = await Effect.runPromise(Xlsx.read(await pkg(onOneId(50), rel("rId1"), sheetData), { limits: { maxDecompressedBytes: 175_000 } }))
      const outcomes: Array<string> = []
      for (let i = 0; i < 50; i++) {
        const exit = await rowsExit(wb, i)
        outcomes.push(Exit.isSuccess(exit) ? "read" : (failureOf(exit) as { limit: string }).limit)
      }
      expect(outcomes).toEqual([...Array(3).fill("read"), ...Array(47).fill("maxDecompressedBytes")])
      // a sheet reading its part again pays nothing new, whatever other sheets spent
      expect(Exit.isSuccess(await rowsExit(wb, 0))).toBe(true)
      expect(Exit.isSuccess(await Effect.runPromise(Effect.exit(Workbook.sheetDetails(wb, 2))))).toBe(true)
      // the 4th sheet still can't, even with the same part already inflated three times
      expect(failureOf(await rowsExit(wb, 3))).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 175_000 })
    })
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

describe("xlsx: inflate budget across reads of one workbook", () => {
  // "Big" inflates to about 140 KB; "Small" to a few hundred bytes
  const bigData = Array.from({ length: 2000 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>${r}</v></c><c r="B${r + 1}"><v>12345</v></c></row>`).join("")
  const book = () => xlsx({ sheets: [{ name: "Small", data: `<row r="1"><c r="A1"><v>1</v></c></row>` }, { name: "Big", data: bigData }] })
  const rowsExit = (wb: Workbook.Workbook, sheet: string, limits?: LimitOverrides) =>
    Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, sheet, limits === undefined ? undefined : { limits }))))
  const detailsExit = (wb: Workbook.Workbook, sheet: string, limits?: LimitOverrides) =>
    Effect.runPromise(Effect.exit(Workbook.sheetDetails(wb, sheet, limits === undefined ? undefined : { limits })))

  it("a sheet already read can be read again after another sheet went over the budget", async () => {
    const wb = await Effect.runPromise(Xlsx.read(await book(), { limits: { maxDecompressedBytes: 4000 } }))
    expect(Exit.isSuccess(await rowsExit(wb, "Small"))).toBe(true)
    expect(failureOf(await rowsExit(wb, "Big"))).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 4000 })
    // re-reading Small spends nothing new
    expect(Exit.isSuccess(await rowsExit(wb, "Small"))).toBe(true)
    expect(Exit.isSuccess(await detailsExit(wb, "Small"))).toBe(true)
    // Big still can't go past the budget
    expect(failureOf(await rowsExit(wb, "Big"))).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 4000 })
  })

  it("inflate limits given to rows and sheetDetails apply, lower or higher than at read time", async () => {
    const bytes = await book()
    // lower: the workbook was read with the defaults
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    const low = { maxDecompressedBytes: 20_000 }
    expect(Exit.isSuccess(await rowsExit(wb, "Small", low))).toBe(true)
    expect(failureOf(await rowsExit(wb, "Big", low))).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 20_000 })
    expect(failureOf(await detailsExit(wb, "Big", low))).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 20_000 })
    const ratio = { maxInflateRatio: 1, inflateFloorBytes: 0 }
    expect(failureOf(await rowsExit(wb, "Big", ratio))).toMatchObject({ _tag: "LimitExceeded", limit: "maxInflateRatio", max: 1 })
    // without overrides the read-time limits apply
    const all = await rowsExit(wb, "Big")
    expect(Exit.isSuccess(all) && all.value.length).toBe(2000)
    // higher: the workbook was read with a budget too small for Big
    const tight = await Effect.runPromise(Xlsx.read(bytes, { limits: { maxDecompressedBytes: 4000 } }))
    expect(failureOf(await detailsExit(tight, "Big"))).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 4000 })
    const raised = await rowsExit(tight, "Big", { maxDecompressedBytes: Infinity })
    expect(Exit.isSuccess(raised) && raised.value.length).toBe(2000)
    expect(Exit.isSuccess(await detailsExit(tight, "Big", { maxDecompressedBytes: Infinity }))).toBe(true)
  })
})

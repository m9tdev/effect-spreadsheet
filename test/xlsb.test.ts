import { readFileSync } from "node:fs"
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { RecordSplitter } from "../src/internal/biff12/records.ts"
import { SheetRows } from "../src/internal/xlsb/sheet.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsb from "../src/xlsb/index.ts"
import { failureOf } from "./helpers/run.ts"
import * as B from "./helpers/xlsb.ts"

const readExit = (bytes: Uint8Array, limits?: LimitOverrides) =>
  Effect.runPromise(Effect.exit(Effect.gen(function*() {
    const wb = yield* Xlsb.read(bytes, { limits })
    const sheets: Array<Array<Workbook.Row>> = []
    const details: Array<Workbook.SheetDetails> = []
    for (const sheet of wb.sheets) {
      sheets.push(Array.from(yield* Stream.runCollect(Workbook.rows(wb, sheet.index))))
      details.push(yield* Workbook.sheetDetails(wb, sheet.index))
    }
    return { wb, sheets, details }
  })))

const readOk = async (book: B.TestBook, limits?: LimitOverrides) => {
  const exit = await readExit(await B.xlsb(book), limits)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value
}
const errorOf = async (book: B.TestBook, limits?: LimitOverrides) => failureOf(await readExit(await B.xlsb(book), limits))
const one = (body: ReadonlyArray<Uint8Array>, book: Omit<B.TestBook, "sheets"> = {}): B.TestBook => ({ ...book, sheets: [{ name: "S", part: B.sheetBin(body) }] })

const values = (rows: ReadonlyArray<Workbook.Row>) =>
  rows.map((r) => [r.index, r.hidden, r.cells.map((c) => [c.column, c._tag, Cell.value(c) instanceof Date ? (Cell.value(c) as Date).toISOString() : Cell.value(c)])])

/** Cell styles for the tests: 0 General, 1 built-in date (14), 2 custom "0.000", 3 custom date. */
const STYLES = { xfs: [0, 14, 164, 165], formats: { 164: "0.000", 165: "yyyy-mm-dd hh:mm" } }

describe("xlsb reader", () => {
  it("reads sheets in order with their visibility, through Spreadsheet.read too", async () => {
    const book: B.TestBook = { sheets: [{ name: "Data" }, { name: "Hidden", state: 1 }, { name: "Very", state: 2 }] }
    const { wb } = await readOk(book)
    expect(wb.format).toBe("xlsb")
    expect(wb.sheets).toEqual([
      { name: "Data", index: 0, visibility: "visible" },
      { name: "Hidden", index: 1, visibility: "hidden" },
      { name: "Very", index: 2, visibility: "veryHidden" }
    ])
    const bytes = await B.xlsb(book)
    expect(await Effect.runPromise(Spreadsheet.detectFormat(bytes))).toBe("xlsb")
    const viaRead = await Effect.runPromise(Spreadsheet.read(bytes))
    expect(viaRead.format).toBe("xlsb")
    expect(viaRead.sheets.map((s) => s.name)).toEqual(["Data", "Hidden", "Very"])
  })

  it("reads every cell type, with numFmt on numbers and dates from date formats", async () => {
    const { sheets } = await readOk(one([
      B.rowHdr(0),
      B.rk(0, B.rkInt(42)),
      B.rk(1, B.rkInt(1234) | 1),
      B.rk(2, B.rkDouble(1.5)),
      B.real(3, 3.25, 2),
      B.cellSt(4, "inline"),
      B.isst(5, 1),
      B.cellBool(6, true),
      B.cellError(7, 0x07),
      B.blank(8),
      B.rstring(9, "rich"),
      B.rowHdr(1),
      B.fmlaNum(0, 7),
      B.fmlaString(1, "result"),
      B.fmlaBool(2, false),
      B.fmlaError(3, 0x2A),
      B.real(4, 45000.5, 1),
      B.rk(5, B.rkInt(45000), 3),
      B.cellError(6, 0x99),
      B.real(7, NaN),
      B.real(8, -0)
    ], { strings: ["zero", "one"], ...STYLES }))
    expect(values(sheets[0]!)).toEqual([
      [0, false, [
        [0, "Number", 42], [1, "Number", 12.34], [2, "Number", 1.5], [3, "Number", 3.25], [4, "Text", "inline"], [5, "Text", "one"],
        [6, "Boolean", true], [7, "Error", "#DIV/0!"], [8, "Empty", undefined], [9, "Text", "rich"]
      ]],
      [1, false, [
        [0, "Number", 7], [1, "Text", "result"], [2, "Boolean", false], [3, "Error", "#N/A"],
        [4, "Date", "2023-03-15T12:00:00.000Z"], [5, "Date", "2023-03-15T00:00:00.000Z"], [6, "Error", "#N/A"], [7, "Error", "#NUM!"], [8, "Number", 0]
      ]]
    ])
    const cells = sheets[0]![0]!.cells
    expect(cells[3]!.numFmt).toBe("0.000")
    expect(cells[0]!.numFmt).toBeUndefined()
    expect(sheets[0]![1]!.cells[4]!.numFmt).toBe("mm-dd-yy")
    expect(sheets[0]![1]!.cells[5]!.numFmt).toBe("yyyy-mm-dd hh:mm")
    expect(Object.is((sheets[0]![1]!.cells[8] as Cell.NumberCell).value, -0)).toBe(false)
    // formulas aren't decompiled
    expect(sheets[0]![1]!.cells.every((c) => c.formula === undefined)).toBe(true)
  })

  it("a cell's style is its low 24 bits: the flags byte above (fPhShow) doesn't change it", async () => {
    const { sheets } = await readOk(one([B.rowHdr(0), B.real(0, 45000, 1 | 0x01000000)], STYLES))
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Date", "2023-03-15T00:00:00.000Z"]]]])
  })

  it("dates in the 1904 system; serials outside Excel's dates stay numbers", async () => {
    const { sheets } = await readOk(one([B.rowHdr(0), B.real(0, 0, 1), B.real(1, -1, 1), B.real(2, 1e9, 1)], { ...STYLES, date1904: true }))
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Date", "1904-01-01T00:00:00.000Z"], [1, "Number", -1], [2, "Number", 1e9]]]])
  })

  it("short cells (no column field) follow the previous cell's column", async () => {
    const { sheets } = await readOk(one([
      B.rowHdr(3),
      B.shortRk(B.rkInt(1)),
      B.shortReal(2.5),
      B.cellSt(5, "five"),
      B.shortSt("six"),
      B.shortIsst(0),
      B.shortBool(true),
      B.shortError(0x17),
      B.shortBlank(),
      B.rowHdr(4),
      B.shortBlank()
    ], { strings: ["s0"] }))
    expect(values(sheets[0]!)).toEqual([
      [3, false, [[0, "Number", 1], [1, "Number", 2.5], [5, "Text", "five"], [6, "Text", "six"], [7, "Text", "s0"], [8, "Boolean", true], [9, "Error", "#REF!"], [10, "Empty", undefined]]],
      [4, false, [[0, "Empty", undefined]]]
    ])
  })

  it("record 0x003F isn't a cell (no short rich string): it adds nothing and doesn't move the short cells' column", async () => {
    // laid out as a short rich string would be: style u32, flags byte, wide string
    const r3f = B.record(0x3F, [...B.u32(0), 1, ...B.wide("not a cell")])
    const { sheets } = await readOk(one([B.rowHdr(0), B.rk(2, B.rkInt(1)), r3f, B.shortRk(B.rkInt(2)), B.rowHdr(1), r3f]))
    expect(values(sheets[0]!)).toEqual([[0, false, [[2, "Number", 1], [3, "Number", 2]]], [1, false, []]])
  })

  it("rows are sparse: rows left out aren't emitted, rows without cells are; far cells cost one cell", async () => {
    const { sheets, details } = await readOk({
      sheets: [{ name: "S", part: B.sheetBin([B.rowHdr(0), B.rk(0, B.rkInt(1)), B.rowHdr(5), B.rowHdr(1048575), B.rk(16383, B.rkInt(2))], { before: [B.wsDim(0, 1048575, 0, 16383)] }) }]
    })
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Number", 1]]], [5, false, []], [1048575, false, [[16383, "Number", 2]]]])
    expect(details[0]!.dimension).toEqual({ firstRow: 0, firstColumn: 0, lastRow: 1048575, lastColumn: 16383 })
  })

  it("cells out of order are sorted, keeping the last of a column", async () => {
    const { sheets } = await readOk(one([B.rowHdr(0), B.rk(3, B.rkInt(3)), B.rk(1, B.rkInt(1)), B.rk(3, B.rkInt(33)), B.rk(0, B.rkInt(0))]))
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Number", 0], [1, "Number", 1], [3, "Number", 33]]]])
  })

  it("the same column twice in a row gives one cell, the last", async () => {
    const { sheets } = await readOk(one([B.rowHdr(0), B.rk(0, B.rkInt(1)), B.rk(0, B.rkInt(2))]))
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Number", 2]]]])
  })

  it("a row repeating its columns holds at most 2 x maxColumns cells", () => {
    const limits = { ...SpreadsheetLimits.defaults, maxColumns: 8 }
    const rows = new SheetRows({ part: "t.bin", strings: undefined, styles: { formats: [] }, date1904: false, limits })
    const splitter = new RecordSplitter("t.bin", limits.maxRecordBytes)
    const feed = (bytes: Uint8Array) => expect(splitter.feed(bytes, rows.handle)).toBeUndefined()
    feed(B.beginSheetData())
    feed(B.rowHdr(0))
    let most = 0
    // progress: k up to 1000
    for (let k = 0; k < 1000; k++) {
      feed(B.rk(7 - (k % 8), B.rkInt(k)))
      most = Math.max(most, rows.heldCells)
    }
    expect(most).toBeLessThanOrEqual(16)
    expect(rows.end()).toBeUndefined()
    expect(rows.out[0]!.cells.map((c) => c.column)).toEqual([0, 1, 2, 3, 4, 5, 6, 7])
    expect(rows.out[0]!.cells.map((c) => Cell.value(c))).toEqual([999, 998, 997, 996, 995, 994, 993, 992])
  })

  it("chart sheets have no rows; macro-enabled packages read the same", async () => {
    const { wb, sheets } = await readOk({
      sheets: [{ name: "S", part: B.sheetBin([B.rowHdr(0), B.rk(0, B.rkInt(1))]) }, { name: "Chart", kind: "chartsheet" }],
      extra: [{ name: "xl/vbaProject.bin", data: new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 1, 2, 3]) }]
    })
    expect(wb.sheets.map((s) => s.name)).toEqual(["S", "Chart"])
    expect(sheets[1]).toEqual([])
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Number", 1]]]])
  })

  it("record numbers inside future-record blocks, and cell numbers outside the sheet data, aren't cells", async () => {
    const frt = [B.record(0x23, [2, 0x0E, 1, 0]), B.record(0x01, [0, 0, 0, 0, 8, 0, 0, 0, 0]), B.record(0x00, [1]), B.record(0x24)]
    const { sheets, details } = await readOk({
      sheets: [{ name: "S", part: B.sheetBin([B.rowHdr(0), B.rk(0, B.rkInt(1)), ...frt, B.rk(1, B.rkInt(2))], { before: frt, after: [...frt, B.blank(3), B.record(0x00, [1])] }) }]
    })
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Number", 1], [1, "Number", 2]]]])
    expect(details[0]).toEqual({ hiddenColumns: [], merges: [] })
  })

  it("stored (uncompressed) parts read the same as deflated ones", async () => {
    const book = one([B.rowHdr(2, true), B.isst(1, 0)], { strings: ["x"], method: "store" })
    const { sheets } = await readOk(book)
    expect(values(sheets[0]!)).toEqual([[2, true, [[1, "Text", "x"]]]])
  })

  it("rows can be streamed again, with other limits", async () => {
    const wb = await Effect.runPromise(Xlsb.read(await B.xlsb(one([B.rowHdr(0), B.rk(0, B.rkInt(1)), B.rowHdr(1), B.rk(0, B.rkInt(2))]))))
    const twice = await Effect.runPromise(Effect.zip(Stream.runCollect(Workbook.rows(wb, 0)), Stream.runCollect(Workbook.rows(wb, "s"))))
    expect(twice[0]).toEqual(twice[1])
    const exit = await Effect.runPromiseExit(Stream.runCollect(Workbook.rows(wb, 0, { limits: { maxRows: 1 } })))
    expect(failureOf(exit)).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows" })
  })
})

describe("xlsb hidden rows (SheetJS 0.20.3's stream reader never ends on these)", () => {
  const sheet = (hidden: ReadonlyArray<number>, rows = 5) =>
    one(Array.from({ length: rows }, (_, r) => [B.rowHdr(r, hidden.includes(r)), B.rk(0, B.rkInt(r))]).flat())
  for (const [label, hidden] of [["first", [0]], ["middle", [2]], ["last", [4]], ["consecutive", [1, 2, 3]], ["all", [0, 1, 2, 3, 4]]] as const) {
    it(`${label}: every row, the hidden ones flagged`, async () => {
      const { sheets } = await readOk(sheet(hidden))
      expect(sheets[0]!.map((r) => [r.index, r.hidden, Cell.value(r.cells[0]!)])).toEqual(
        [0, 1, 2, 3, 4].map((r) => [r, hidden.includes(r as never), r])
      )
    })
  }

  it("a hidden row without cells is emitted too", async () => {
    const { sheets } = await readOk(one([B.rowHdr(0, true), B.rowHdr(1), B.rk(0, B.rkInt(1))]))
    expect(values(sheets[0]!)).toEqual([[0, true, []], [1, false, [[0, "Number", 1]]]])
  })

  it("the committed hidden-row fixtures read with every row", async () => {
    for (const path of ["test/corpus/hidden-rows/hidden-row.xlsb", "test/corpus/generated/hidden_rows.xlsb"]) {
      const exit = await readExit(new Uint8Array(readFileSync(path)))
      if (Exit.isFailure(exit)) throw new Error(`${path}: ${JSON.stringify(failureOf(exit))}`)
      const rows = exit.value.sheets[0]!
      expect(rows.some((r) => r.hidden), path).toBe(true)
      expect(rows.map((r) => r.index), path).toEqual(rows.map((_, k) => rows[0]!.index + k))
    }
  })
})

describe("xlsb sheet details", () => {
  it("hidden columns as merged spans, merges as listed, the declared dimension", async () => {
    const { details } = await readOk({
      sheets: [{
        name: "S",
        part: B.sheetBin([B.rowHdr(0), B.rk(0, B.rkInt(1))], {
          before: [B.wsDim(0, 9, 0, 4), B.beginColInfos(), B.colInfo(1, 1, true), B.colInfo(2, 3, true), B.colInfo(4, 4, false), B.colInfo(6, 20000, true), B.endColInfos()],
          after: [B.beginMergeCells(2), B.mergeCell(0, 1, 0, 1), B.mergeCell(5, 5, 2, 4), B.endMergeCells()]
        })
      }]
    })
    expect(details[0]).toEqual({
      hiddenColumns: [{ first: 1, last: 3 }, { first: 6, last: 16383 }],
      merges: [{ firstRow: 0, lastRow: 1, firstColumn: 0, lastColumn: 1 }, { firstRow: 5, lastRow: 5, firstColumn: 2, lastColumn: 4 }],
      dimension: { firstRow: 0, firstColumn: 0, lastRow: 9, lastColumn: 4 }
    })
  })

  it("no BrtWsDim: no dimension; a reversed one is ignored", async () => {
    const plain = await readOk(one([B.rowHdr(0)]))
    expect(plain.details[0]).toEqual({ hiddenColumns: [], merges: [] })
    const reversed = await readOk({ sheets: [{ name: "S", part: B.sheetBin([], { before: [B.wsDim(5, 1, 0, 0)] }) }] })
    expect(reversed.details[0]!.dimension).toBeUndefined()
  })

  it("column infos past maxColumns, and a bad merge, fail", async () => {
    const flood = { sheets: [{ name: "S", part: B.sheetBin([], { before: Array.from({ length: 9 }, (_, k) => B.colInfo(k, k, true)) }) }] }
    expect(await errorOf(flood, { maxColumns: 8 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxColumns", actual: 9 })
    const bad = { sheets: [{ name: "S", part: B.sheetBin([], { after: [B.mergeCell(3, 1, 0, 0)] }) }] }
    expect(await errorOf(bad)).toMatchObject({ _tag: "MalformedFile", format: "xlsb", reason: "bad merge range" })
  })
})

describe("xlsb defined names", () => {
  const sheets = [{ name: "Data" }, { name: "My Sheet" }, { name: "Jan" }, { name: "Mar" }]
  const globals = [
    B.record(0x161),
    B.supSelf(),
    B.supBookSrc("rIdX1"),
    B.supAddin(),
    B.supBookSrc("rIdMissing"),
    // XTIs: 0 Data, 1 My Sheet, 2 Jan:Mar, 3 deleted sheet, 4 external sheet 1, 5 add-in, 6 external without a part, 7 out of range
    B.externSheet([[0, 0, 0], [0, 1, 1], [0, 2, 3], [0, -1, -1], [1, 1, 1], [2, 0, 0], [3, 0, 0], [0, 9, 9]]),
    B.record(0x162)
  ]
  const external = B.concat([B.record(0x168, [...B.u16(0), ...B.wide("rId1"), ...B.u32(0xFFFFFFFF)]), B.supTabs(["Ext0", "Rates"])])

  it("single 3-D references and areas are written as Excel writes them in XLSX", async () => {
    const { wb } = await readOk({
      sheets,
      externalLinks: [external],
      globals: [
        ...globals,
        B.name("Cell", B.ref3d(0, 1, 2)),
        B.name("Relative", B.ref3d(0, 1, 2 | 0xC000)),
        B.name("Area", B.area3d(1, 0, 9, 0, 2)),
        B.name("Rows", B.area3d(0, 0, 1, 0, 0x3FFF)),
        B.name("Columns", B.area3d(0, 0, 0xFFFFF, 0, 2)),
        B.name("Span", B.area3d(2, 0, 9, 0, 0)),
        B.name("Deleted", B.ref3d(3, 0, 0)),
        B.name("Rate", B.ref3d(4, 0, 0)),
        B.name("AddIn", B.ref3d(5, 0, 0)),
        B.name("NoLink", B.ref3d(6, 0, 0)),
        B.name("Outside", B.ref3d(7, 0, 0)),
        B.name("NoXti", B.ref3d(99, 0, 0)),
        B.name("Lost", B.refErr3d(0)),
        B.name("LostArea", B.areaErr3d(3)),
        B.name("Calc", [0x1E, 42, 0]),
        B.name("Empty", []),
        B.name("Print_Area", B.area3d(0, 0, 4, 0, 3), { builtin: true, itab: 0 }),
        B.name("Scoped", B.ref3d(0, 0, 0), { itab: 1, hidden: true }),
        B.name("BadScope", B.ref3d(0, 0, 0), { itab: 7 })
      ]
    })
    expect(wb.definedNames).toEqual([
      { name: "Cell", formula: "Data!$C$2", hidden: false },
      { name: "Relative", formula: "Data!C2", hidden: false },
      { name: "Area", formula: "'My Sheet'!$A$1:$C$10", hidden: false },
      { name: "Rows", formula: "Data!$1:$2", hidden: false },
      { name: "Columns", formula: "Data!$A:$C", hidden: false },
      { name: "Span", formula: "Jan:Mar!$A$1:$A$10", hidden: false },
      { name: "Deleted", formula: "#REF!$A$1", hidden: false },
      { name: "Rate", formula: "[1]Rates!$A$1", hidden: false },
      { name: "AddIn", formula: "#REF!$A$1", hidden: false },
      { name: "NoLink", formula: "#REF!$A$1", hidden: false },
      { name: "Outside", formula: "#REF!$A$1", hidden: false },
      { name: "NoXti", formula: "", hidden: false },
      { name: "Lost", formula: "Data!#REF!", hidden: false },
      { name: "LostArea", formula: "#REF!", hidden: false },
      { name: "Calc", formula: "", hidden: false },
      { name: "Empty", formula: "", hidden: false },
      { name: "_xlnm.Print_Area", formula: "Data!$A$1:$D$5", scope: 0, hidden: false },
      { name: "Scoped", formula: "Data!$A$1", scope: 1, hidden: true },
      { name: "BadScope", formula: "Data!$A$1", hidden: false }
    ])
  })

  it("a union of two areas (print titles' form) isn't decompiled: \"\"", async () => {
    const union = [...B.area3d(0, 0, 0xFFFFF, 0, 1), ...B.area3d(0, 0, 1, 0, 0x3FFF), 0x10]
    const { wb } = await readOk({ sheets, globals: [B.supSelf(), B.externSheet([[0, 0, 0]]), B.name("Print_Titles", union, { itab: 0, builtin: true })] })
    expect(wb.definedNames).toEqual([{ name: "_xlnm.Print_Titles", formula: "", scope: 0, hidden: false }])
  })

  it("without supporting books, XTIs of book 0 mean this workbook", async () => {
    const { wb } = await readOk({ sheets, globals: [B.externSheet([[0, 1, 1]]), B.name("N", B.ref3d(0, 0, 0))] })
    expect(wb.definedNames).toEqual([{ name: "N", formula: "'My Sheet'!$A$1", hidden: false }])
  })

  it("names over maxDefinedNames, a truncated name or external sheet record fail", async () => {
    const names = Array.from({ length: 4 }, (_, k) => B.name(`N${k}`, []))
    expect(await errorOf({ sheets, globals: names }, { maxDefinedNames: 3 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxDefinedNames", actual: 4 })
    expect(await errorOf({ sheets, globals: [B.record(0x27, [0, 0, 0, 0, 0, 0xFF, 0xFF, 0xFF, 0xFF, ...B.u32(5), 65, 0])] })).toMatchObject({ _tag: "MalformedFile", reason: "truncated defined name" })
    const formulaPast = B.record(0x27, [0, 0, 0, 0, 0, 0xFF, 0xFF, 0xFF, 0xFF, ...B.wide("N"), ...B.u32(100), 0x3A])
    expect(await errorOf({ sheets, globals: [formulaPast] })).toMatchObject({ _tag: "MalformedFile", reason: "defined name formula past its record" })
    expect(await errorOf({ sheets, globals: [B.record(0x16A, [...B.u32(2), ...B.u32(0), ...B.u32(0), ...B.u32(0)])] })).toMatchObject({ _tag: "MalformedFile", reason: "truncated external sheet record" })
  })

  it("an external link's sheet list over maxSheets fails", async () => {
    const big = B.concat([B.supTabs(["a", "b", "c", "d", "e"])])
    expect(await errorOf({ sheets: [{ name: "S" }], externalLinks: [big], globals: [B.supSelf(), B.supBookSrc("rIdX1"), B.externSheet([[1, 0, 0]]), B.name("N", B.ref3d(0, 0, 0))] }, { maxSheets: 4 }))
      .toMatchObject({ _tag: "LimitExceeded", limit: "maxSheets", actual: 5 })
  })
})

describe("xlsb limits and malformed input", () => {
  it("row and column indexes, rows, cells and strings are limited", async () => {
    expect(await errorOf(one([B.rowHdr(10)]), { maxRows: 10 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", actual: 11 })
    expect(await errorOf(one([B.rowHdr(0), B.rowHdr(0), B.rowHdr(0)]), { maxRows: 2 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", actual: 3 })
    expect(await errorOf(one([B.rowHdr(0), B.rk(8, 0)]), { maxColumns: 8 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxColumns", actual: 9 })
    expect(await errorOf(one([B.rowHdr(0), B.blank(0), B.blank(1), B.blank(2)]), { maxCellsPerSheet: 2 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxCellsPerSheet", actual: 3 })
    expect(await errorOf(one([B.rowHdr(0), B.cellSt(0, "abcde")]), { maxStringLength: 4 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
    expect(await errorOf(one([], { strings: ["abcde"] }), { maxStringLength: 4 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
    expect(await errorOf(one([], { strings: ["a", "b", "c"] }), { maxSharedStrings: 2 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxSharedStrings", actual: 3 })
    expect(await errorOf({ sheets: [{ name: "a" }, { name: "b" }, { name: "c" }] }, { maxSheets: 2 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxSheets", actual: 3 })
    expect(await errorOf(one([B.rowHdr(0), B.cellSt(0, "x".repeat(100))]), { maxRecordBytes: 64 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxRecordBytes" })
  })

  it("broken records are MalformedFile at their offset", async () => {
    expect(await errorOf(one([B.rk(0, 1)]))).toMatchObject({ _tag: "MalformedFile", format: "xlsb", part: "xl/worksheets/sheet1.bin", reason: "cell record 0x2 outside a row" })
    expect(await errorOf(one([B.record(0, [1, 0, 0, 0])]))).toMatchObject({ _tag: "MalformedFile", reason: "truncated row header" })
    expect(await errorOf(one([B.rowHdr(0), B.record(5, [0, 0, 0, 0, 0, 0, 0, 0, 1, 2])]))).toMatchObject({ _tag: "MalformedFile", reason: "truncated cell record 0x5" })
    expect(await errorOf(one([B.rowHdr(0), B.isst(0, 2)], { strings: ["a"] }))).toMatchObject({ _tag: "MalformedFile", reason: "shared string 2 doesn't exist" })
    expect(await errorOf(one([B.rowHdr(0), B.isst(0, 0)]))).toMatchObject({ _tag: "MalformedFile", reason: "shared string 0 doesn't exist" })
    const cut = B.sheetBin([B.rowHdr(0), B.rk(0, 1)])
    expect(await errorOf({ sheets: [{ name: "S", part: cut.subarray(0, cut.length - 8) }] })).toMatchObject({ _tag: "MalformedFile", reason: expect.stringContaining("runs past the end of the part") })
  })

  it("a broken package: no workbook part, a part that isn't a workbook, a missing sheet part, two sheets on one part", async () => {
    expect(await errorOf({ sheets: [], workbookPart: new Uint8Array(0) })).toMatchObject({ _tag: "MalformedFile", reason: "empty workbook part" })
    expect(await errorOf({ sheets: [], workbookPart: B.concat([B.record(0x81)]) })).toMatchObject({ _tag: "MalformedFile", reason: "not a workbook part: it starts with record 0x81" })
    expect(await errorOf({ sheets: [{ name: "S" }], workbookPart: B.concat([B.beginBook(), B.bundleSh("S", "rId9")]) })).toMatchObject({ _tag: "MalformedFile", reason: "sheet \"S\" has no part" })
    const shared = B.concat([B.beginBook(), B.bundleSh("A", "rId1"), B.bundleSh("B", "rId1")])
    expect(await errorOf({ sheets: [{ name: "A" }], workbookPart: shared })).toMatchObject({ _tag: "MalformedFile", reason: "sheets \"A\" and \"B\" share part xl/worksheets/sheet1.bin" })
    const truncatedSheet = B.concat([B.beginBook(), B.record(0x9C, [0, 0, 0, 0, 1, 0, 0, 0, 0xFF, 0xFF, 0xFF, 0xFF, 5, 0, 0, 0, 65, 0])])
    expect(await errorOf({ sheets: [], workbookPart: truncatedSheet })).toMatchObject({ _tag: "MalformedFile", reason: "truncated sheet record" })
  })
})

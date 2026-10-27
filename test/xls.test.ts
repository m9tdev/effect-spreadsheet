import { readFileSync } from "node:fs"
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { codepageLabel, decodeRk } from "../src/internal/xls/fields.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { writeCfb } from "./helpers/cfb.ts"
import { failureOf } from "./helpers/run.ts"
import * as X from "./helpers/xls.ts"

const readExit = (bytes: Uint8Array, limits?: LimitOverrides) =>
  Effect.runPromise(Effect.exit(Effect.gen(function*() {
    const wb = yield* Xls.read(bytes, { limits })
    const sheets: Array<Array<Workbook.Row>> = []
    const details: Array<Workbook.SheetDetails> = []
    for (const sheet of wb.sheets) {
      sheets.push(Array.from(yield* Stream.runCollect(Workbook.rows(wb, sheet.index))))
      details.push(yield* Workbook.sheetDetails(wb, sheet.index))
    }
    return { wb, sheets, details }
  })))

const readOk = async (book: X.TestBook, limits?: LimitOverrides) => {
  const exit = await readExit(X.xls(book), limits)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value
}
const errorOf = async (book: X.TestBook, limits?: LimitOverrides) => failureOf(await readExit(X.xls(book), limits))
const one = (records: ReadonlyArray<Uint8Array>, globals: ReadonlyArray<Uint8Array> = [], biff: 5 | 8 = 8): X.TestBook => ({ biff, globals, sheets: [{ name: "S", records }] })

const values = (rows: ReadonlyArray<Workbook.Row>) =>
  rows.map((r) => [r.index, r.hidden, r.cells.map((c) => [c.column, c._tag, Cell.value(c) instanceof Date ? (Cell.value(c) as Date).toISOString() : Cell.value(c)])])

/** Cell styles for the tests: XF 0 General, 1 built-in date (14), 2 custom "0.000", 3 custom date. */
const STYLES = [X.format(164, "0.000"), X.format(165, "yyyy-mm-dd hh:mm"), X.xf(0), X.xf(14), X.xf(164), X.xf(165)]

describe("xls fields", () => {
  it("RK numbers: integers, doubles, divided by 100", () => {
    expect(decodeRk(X.rkInt(42))).toBe(42)
    expect(decodeRk(X.rkInt(-7))).toBe(-7)
    expect(decodeRk(X.rkInt(1234) | 1)).toBe(12.34)
    expect(decodeRk(X.rkDouble(1.5))).toBe(1.5)
    expect(decodeRk(X.rkDouble(-2.25))).toBe(-2.25)
    expect(decodeRk(X.rkDouble(100) | 1)).toBe(1)
  })

  it("code pages map to TextDecoder labels, unknown ones to windows-1252", () => {
    expect([1252, 1251, 1200, 932, 10000, 65001, 28592, 437, 850].map(codepageLabel)).toEqual([
      "windows-1252", "windows-1251", "utf-16le", "shift_jis", "macintosh", "utf-8", "iso-8859-2", "windows-1252", "windows-1252"
    ])
    for (const label of ["windows-1251", "shift_jis", "macintosh", "iso-8859-2", "ibm866", "koi8-r", "gbk", "big5", "euc-kr", "windows-874"]) {
      expect(() => new TextDecoder(label), label).not.toThrow()
    }
  })
})

describe("xls: cells and rows (BIFF8)", () => {
  it("every cell record type, with number formats and dates", async () => {
    const { sheets } = await readOk(one([
      X.row(0), X.row(1),
      X.labelSst(0, 0, 0, 1), X.label(0, 1, 0, "ünïcödé 名前"), X.label(0, 2, 0, "plain"), X.number(0, 3, 2, 2.5),
      X.rk(0, 4, 0, X.rkInt(-7)), X.rk(0, 5, 0, X.rkInt(1234) | 1), X.mulrk(0, 6, [[0, X.rkInt(1)], [0, X.rkDouble(1.5)]]),
      X.boolErr(1, 0, 0, 1, false), X.boolErr(1, 1, 0, 0x07, true), X.blank(1, 2, 0), X.mulBlank(1, 3, [0, 0]),
      X.number(1, 5, 3, 46024.5208333333), X.rk(1, 6, 1, X.rkInt(61)), X.number(1, 7, 1, 60), X.number(1, 8, 1, -1)
    ], [...STYLES, X.sst(["zero", "one"])]))
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Text", "one"], [1, "Text", "ünïcödé 名前"], [2, "Text", "plain"], [3, "Number", 2.5], [4, "Number", -7], [5, "Number", 12.34], [6, "Number", 1], [7, "Number", 1.5]]],
      [1, false, [[0, "Boolean", true], [1, "Error", "#DIV/0!"], [2, "Empty", undefined], [3, "Empty", undefined], [4, "Empty", undefined],
        [5, "Date", "2026-01-02T12:30:00.000Z"], [6, "Date", "1900-03-01T00:00:00.000Z"], [7, "Number", 60], [8, "Number", -1]]]
    ])
    const cells = sheets[0]!.flatMap((r) => r.cells)
    expect(cells.find((c) => c.column === 3)?.numFmt).toBe("0.000")
    expect(sheets[0]![1]!.cells[5]!.numFmt).toBe("yyyy-mm-dd hh:mm")
    expect(sheets[0]![1]!.cells[6]!.numFmt).toBe("mm-dd-yy")
  })

  it("formula results: number, string (from STRING), boolean, error, empty string", async () => {
    const { sheets } = await readOk(one([
      X.formula(0, 0, 0, 3.5), X.formula(0, 1, 0, { kind: 0 }), X.string("result"), X.formula(0, 2, 0, { kind: 1, value: 1 }),
      X.formula(0, 3, 0, { kind: 2, value: 0x2A }), X.formula(0, 4, 0, { kind: 3 }), X.formula(0, 5, 0, { kind: 0 }), X.number(0, 6, 0, 9)
    ]))
    expect(values(sheets[0]!)).toEqual([[0, false, [
      [0, "Number", 3.5], [1, "Text", "result"], [2, "Boolean", true], [3, "Error", "#N/A"], [4, "Text", ""], [5, "Text", ""], [6, "Number", 9]
    ]]])
  })

  it("rows: hidden ROW records, rows without cells, out-of-order cells, a later row block", async () => {
    const { sheets } = await readOk(one([
      X.row(0), X.row(1, true), X.row(2, true), X.row(3),
      X.number(0, 2, 0, 3), X.number(0, 0, 0, 1), X.number(0, 2, 0, 33),
      X.number(2, 0, 0, 5),
      X.row(40, true), X.number(40, 1, 0, 6)
    ]))
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Number", 1], [2, "Number", 33]]],
      [1, true, []],
      [2, true, [[0, "Number", 5]]],
      [3, false, []],
      [40, true, [[1, "Number", 6]]]
    ])
  })

  it("the 1904 date system, and an embedded chart substream inside a sheet", async () => {
    const { sheets } = await readOk(one([
      X.number(0, 0, 1, 0),
      X.bof(0x0020), X.number(5, 5, 0, 99), X.eof(), // a chart object's substream: skipped
      X.number(1, 0, 1, 60)
    ], [...STYLES, X.datemode1904()]))
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Date", "1904-01-01T00:00:00.000Z"]]],
      [1, false, [[0, "Date", "1904-03-01T00:00:00.000Z"]]]
    ])
  })

  it("sheet visibility and kinds: chart, macro and VB module sheets", async () => {
    const { wb, sheets } = await readOk({
      sheets: [
        { name: "Data", records: [X.number(0, 0, 0, 1)] },
        { name: "Hid", state: 1, records: [] },
        { name: "Very", state: 2, records: [] },
        { name: "Chart", dt: 2, kind: 0x0020, records: [X.number(0, 0, 0, 9)] },
        { name: "Macro", dt: 1, kind: 0x0040, records: [X.number(0, 0, 0, 9)] },
        { name: "Module1", dt: 6, kind: 0x0006, records: [] }
      ]
    })
    expect(wb.sheets.map((s) => [s.name, s.visibility])).toEqual([["Data", "visible"], ["Hid", "hidden"], ["Very", "veryHidden"], ["Chart", "visible"], ["Macro", "visible"]])
    expect(sheets.map((r) => r.length)).toEqual([1, 0, 0, 0, 0])
  })

  it("sheetDetails: hidden columns (merged spans), merges, the dimension", async () => {
    const { details } = await readOk(one([
      X.dimensions(1, 9, 0, 4), X.colInfo(10, 12, true), X.colInfo(1, 2, true), X.colInfo(3, 3, true), X.colInfo(5, 5, false),
      X.number(1, 0, 0, 1), X.mergeCells([[0, 0, 0, 2], [1, 2, 0, 0]]),
      X.bof(0x0020), X.mergeCells([[9, 9, 9, 9]]), X.eof()
    ]))
    expect(details[0]).toEqual({
      hiddenColumns: [{ first: 1, last: 3 }, { first: 10, last: 12 }],
      merges: [{ firstRow: 0, firstColumn: 0, lastRow: 0, lastColumn: 2 }, { firstRow: 1, firstColumn: 0, lastRow: 2, lastColumn: 0 }],
      dimension: { firstRow: 1, firstColumn: 0, lastRow: 8, lastColumn: 3 }
    })
  })

  it("defined names: 3-D references and areas, built-in names, scope, hidden; other formulas are empty", async () => {
    const { wb } = await readOk({
      globals: [
        X.externSheet([0, 1]),
        X.name(0x06, X.area3d(1, 0, 9, 0, 2), { itab: 2, hidden: true }),
        X.name("Total", X.ref3d(0, 4, 1)),
        X.name("Rel", X.ref3d(0, 4, 0xC001)),
        X.name("Sum", [0x1E, 1, 0, 0x1E, 2, 0, 0x03]),
        X.name("Far", X.ref3d(7, 0, 0))
      ],
      sheets: [{ name: "Data", records: [] }, { name: "My Sheet", records: [] }]
    })
    expect(wb.definedNames).toEqual([
      { name: "_xlnm.Print_Area", formula: "'My Sheet'!$A$1:$C$10", hidden: true, scope: 1 },
      { name: "Total", formula: "Data!$B$5", hidden: false },
      { name: "Rel", formula: "Data!B5", hidden: false },
      { name: "Sum", formula: "", hidden: false },
      { name: "Far", formula: "", hidden: false }
    ])
  })

  it("rows can be streamed twice; limits from read time stay with the workbook", async () => {
    const bytes = X.xls(one([X.number(0, 0, 0, 1), X.number(1, 0, 0, 2), X.number(2, 0, 0, 3)]))
    const wb = await Effect.runPromise(Xls.read(bytes).pipe(Effect.provide(SpreadsheetLimits.layer({ maxRows: 2 }))))
    expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)))))).toMatchObject({ limit: "maxRows", max: 2 })
    const rows = Workbook.rows(wb, 0, { limits: { maxRows: 3 } })
    expect(Array.from(await Effect.runPromise(Stream.runCollect(rows))).length).toBe(3)
    expect(Array.from(await Effect.runPromise(Stream.runCollect(rows))).length).toBe(3)
  })
})

describe("xls: BIFF5, other containers, BIFF2-4", () => {
  it("BIFF5: a Book stream, code-page strings for sheet names, labels and formats", async () => {
    const cyrillic = [0xCF, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2] // "Привет" in windows-1251
    const { wb, sheets } = await readOk({
      biff: 5,
      globals: [X.codepage(1251), X.format(164, "0.0", 5), X.xf(0), X.xf(164)],
      sheets: [{ name: "x", name5: cyrillic, records: [X.label5(0, 0, 0, cyrillic), X.number(0, 1, 1, 2), X.dimensions(0, 1, 0, 2, 5)] }]
    })
    expect(wb.sheets[0]!.name).toBe("Привет")
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Text", "Привет"], [1, "Number", 2]]]])
    expect(sheets[0]![0]!.cells[1]!.numFmt).toBe("0.0")
    const details = await Effect.runPromise(Effect.flatMap(Xls.read(X.xls({ biff: 5, sheets: [{ name: "x", records: [X.dimensions(0, 1, 0, 2, 5)] }] })), (w) => Workbook.sheetDetails(w, 0)))
    expect(details.dimension).toEqual({ firstRow: 0, firstColumn: 0, lastRow: 0, lastColumn: 1 })
  })

  it("a bare BIFF8 workbook stream (no compound file) reads", async () => {
    const exit = await readExit(X.workbookStream(one([X.number(0, 0, 0, 7)])))
    expect(Exit.isSuccess(exit) && values(exit.value.sheets[0]!)).toEqual([[0, false, [[0, "Number", 7]]]])
  })

  it("BIFF2, BIFF3 and BIFF4 files fail with UnsupportedFormat naming the version", async () => {
    for (const v of [2, 3, 4]) {
      const bytes = new Uint8Array(readFileSync(`test/corpus/generated/basic.biff${v}.xls`))
      for (const run of [Xls.read(bytes), Spreadsheet.read(bytes)]) {
        expect(failureOf(await Effect.runPromise(Effect.exit(run)))).toMatchObject({ _tag: "UnsupportedFormat", format: `biff${v}` })
      }
    }
  })

  it("password-protected and encrypted files are Encrypted; a compound file without a workbook is UnsupportedFormat", async () => {
    expect(await errorOf({ globals: [X.filepass()], sheets: [{ name: "S", records: [] }] })).toMatchObject({ _tag: "Encrypted", format: "xls" })
    const encrypted = writeCfb([{ name: "EncryptionInfo", data: new Uint8Array(200) }, { name: "EncryptedPackage", data: new Uint8Array(5000) }]).bytes
    for (const run of [Xls.read(encrypted), Spreadsheet.read(encrypted)]) {
      expect(failureOf(await Effect.runPromise(Effect.exit(run)))).toMatchObject({ _tag: "Encrypted", format: "xlsx" })
    }
    const other = writeCfb([{ name: "WordDocument", data: new Uint8Array(100) }]).bytes
    expect(failureOf(await readExit(other))).toMatchObject({ _tag: "UnsupportedFormat", format: "cfb" })
    expect(await Effect.runPromise(Spreadsheet.detectFormat(X.xls(one([]))))).toBe("xls")
    expect(await Effect.runPromise(Spreadsheet.detectFormat(other))).toBe("cfb")
  })
})

describe("xls: hostile and broken files", () => {
  it("bad structure is MalformedFile", async () => {
    const bad = async (stream: Uint8Array) => failureOf(await readExit(writeCfb([{ name: "Workbook", data: stream }]).bytes))
    expect(await bad(new Uint8Array(0))).toMatchObject({ _tag: "MalformedFile", reason: "empty workbook stream" })
    expect(await bad(X.number(0, 0, 0, 1))).toMatchObject({ reason: "the workbook stream doesn't start with a workbook BOF" })
    expect(await bad(X.bof(0x0005))).toMatchObject({ reason: "the workbook globals have no EOF" })
    // a sheet offset past the stream, and one not at a BOF
    const stream = X.workbookStream(one([X.number(0, 0, 0, 1)]))
    const view = new DataView(stream.buffer)
    const sheetAt = stream.indexOf(0x85) // the BOUNDSHEET record's type
    const far = stream.slice()
    new DataView(far.buffer).setUint32(sheetAt + 4, 1_000_000, true)
    const farWb = await Effect.runPromise(Xls.read(writeCfb([{ name: "Workbook", data: far }]).bytes))
    expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(farWb, 0)))))).toMatchObject({ reason: "sheet offset 1000000 is outside the workbook stream" })
    const off = stream.slice()
    new DataView(off.buffer).setUint32(sheetAt + 4, view.getUint32(sheetAt + 4, true) + 4, true)
    const offWb = await Effect.runPromise(Xls.read(writeCfb([{ name: "Workbook", data: off }]).bytes))
    expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(offWb, 0)))))).toMatchObject({ _tag: "MalformedFile" })
  })

  it("bad records are MalformedFile", async () => {
    const cases: Array<[string, Uint8Array]> = [
      ["truncated record 0x203", X.record(0x0203, [0, 0, 0, 0, 0, 0, 1, 2])],
      ["shared string 5 doesn't exist", X.labelSst(0, 0, 0, 5)],
      ["record 0xbd doesn't match its column range", X.record(0x00BD, [...X.u16(0), ...X.u16(0), ...X.u16(0), ...X.u32(1), ...X.u16(9)])]
    ]
    for (const [reason, cell] of cases) {
      expect(await errorOf(one([cell], [X.sst(["only"])])), reason).toMatchObject({ _tag: "MalformedFile", reason })
    }
    // a sheet without its EOF: the record walk ends at the stream
    const stream = X.workbookStream(one([X.number(0, 0, 0, 1)]))
    const cut = stream.subarray(0, stream.length - 4)
    const wb = await Effect.runPromise(Xls.read(writeCfb([{ name: "Workbook", data: cut }]).bytes))
    expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)))))).toMatchObject({ reason: "the sheet has no EOF" })
  })

  it("limits: rows, columns, cells, sheets, shared strings, string lengths", async () => {
    expect(await errorOf(one([X.number(5, 0, 0, 1)]), { maxRows: 5 })).toMatchObject({ limit: "maxRows", actual: 6 })
    expect(await errorOf(one([X.number(0, 300, 0, 1)]), { maxColumns: 256 })).toMatchObject({ limit: "maxColumns", actual: 301 })
    expect(await errorOf(one([X.mulBlank(0, 0, [0, 0, 0])]), { maxCellsPerSheet: 2 })).toMatchObject({ limit: "maxCellsPerSheet" })
    expect(await errorOf({ sheets: [{ name: "a", records: [] }, { name: "b", records: [] }] }, { maxSheets: 1 })).toMatchObject({ limit: "maxSheets" })
    expect(await errorOf(one([], [X.sst(["a", "b", "c"])]), { maxSharedStrings: 2 })).toMatchObject({ limit: "maxSharedStrings", actual: 3 })
    expect(await errorOf(one([], [X.sst(["x".repeat(50)])]), { maxStringLength: 10 })).toMatchObject({ limit: "maxStringLength" })
    expect(await errorOf(one([X.label(0, 0, 0, "y".repeat(50))]), { maxStringLength: 10 })).toMatchObject({ limit: "maxStringLength" })
    expect(await errorOf(one([X.colInfo(0, 0, true), X.colInfo(1, 1, true)]), { maxColumns: 1 })).toMatchObject({ limit: "maxColumns" })
    expect(await errorOf(one([X.mergeCells([[0, 0, 0, 1], [1, 1, 0, 1]])]), { maxCellsPerSheet: 1 })).toMatchObject({ limit: "maxCellsPerSheet", actual: 2 })
    expect(await errorOf(one([X.label5(0, 0, 0, new Array(50).fill(0x41))], [], 5), { maxStringLength: 10 })).toMatchObject({ limit: "maxStringLength" })
    // a formula's string result (STRING) and a sheet name (BOUNDSHEET)
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), X.string("z".repeat(50))]), { maxStringLength: 10 })).toMatchObject({ limit: "maxStringLength" })
    for (const biff of [8, 5] as const) {
      expect(await errorOf({ biff, sheets: [{ name: "n".repeat(50), records: [] }] }, { maxStringLength: 10 })).toMatchObject({ limit: "maxStringLength" })
    }
  })

  it("a hidden column span past the last column is clamped to maxColumns", async () => {
    const { details } = await readOk(one([X.colInfo(2, 0xFFFF, true)]), { maxColumns: 256 })
    expect(details[0]!.hiddenColumns).toEqual([{ first: 2, last: 255 }])
  })

  it("an SST split across CONTINUE records reads, the strings' widths switching at the boundary", async () => {
    // "abc" then "日本" split: "日" in the SST record, "本" in a CONTINUE with its option byte
    const head = [...X.u32(2), ...X.u32(2), ...X.xlString("abc"), ...X.u16(2), 1, ...X.u16("日".charCodeAt(0))]
    const sstRecords = [X.record(0x00FC, head), X.record(0x003C, [1, ...X.u16("本".charCodeAt(0))])]
    const { sheets } = await readOk(one([X.labelSst(0, 0, 0, 0), X.labelSst(0, 1, 0, 1)], sstRecords))
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Text", "abc"], [1, "Text", "日本"]]]])
  })
})

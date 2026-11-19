/**
 * XLS worksheet substreams (src/internal/xls/sheet.ts): formula string
 * results across SHRFMLA/ARRAY/TABLE and CONTINUE, the BIFF4 FORMULA record
 * number, rows whose cells come in several places, non-finite numbers and
 * error codes, and sheetDetails edge cases. Synthetic records only.
 */
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it, vi } from "vitest"
import * as Cell from "../src/Cell.ts"
import { type Meter, noMeter } from "../src/internal/meter.ts"
import { readGlobals } from "../src/internal/xls/globals.ts"
import { BATCH_RECORDS, DISPLACED_RECORDS, sheetRows } from "../src/internal/xls/sheet.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { writeCfb } from "./helpers/cfb.ts"
import { counter, metered } from "./helpers/meter.ts"
import { failureOf } from "./helpers/run.ts"
import { countYields } from "./helpers/yields.ts"
import * as X from "./helpers/xls.ts"

// A reader that stops advancing (e.g. a broken `behind`) loops with yields
// instead of ending: reads time out in seconds, and tests in 10, rather than
// each waiting for the run's 60-second test timeout.
vi.setConfig({ testTimeout: 10_000 })

const readExit = (bytes: Uint8Array, limits?: LimitOverrides) =>
  Effect.runPromise(Effect.exit(Effect.timeout(Effect.gen(function*() {
    const wb = yield* Xls.read(bytes, { limits })
    const sheets: Array<Array<Workbook.Row>> = []
    const details: Array<Workbook.SheetDetails> = []
    for (const sheet of wb.sheets) {
      sheets.push(Array.from(yield* Stream.runCollect(Workbook.rows(wb, sheet.index))))
      details.push(yield* Workbook.sheetDetails(wb, sheet.index))
    }
    return { wb, sheets, details }
  }), "5 seconds")))

const readOk = async (book: X.TestBook, limits?: LimitOverrides) => {
  const exit = await readExit(X.xls(book), limits)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value
}
const errorOf = async (book: X.TestBook, limits?: LimitOverrides) => failureOf(await readExit(X.xls(book), limits))
/** The rows of a book's first sheet, read with `displacedRecords` as the cap on records out of place. */
const rowsDisplacing = (book: X.TestBook, displacedRecords: number) =>
  Effect.runPromise(Effect.exit(Effect.gen(function*() {
    const stream = X.workbookStream(book)
    const limits = SpreadsheetLimits.defaults
    const globals = yield* readGlobals(stream, limits)
    const { offset, end } = globals.sheets[0]!
    return Array.from(yield* Stream.runCollect(sheetRows({ stream, offset, end, part: "sheet", globals, limits, displacedRecords })))
  })))
const rowsDisplaced = async (book: X.TestBook, displacedRecords: number) => {
  const exit = await rowsDisplacing(book, displacedRecords)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value
}
/** The rows a sheet's stream gives before it fails (or all of them), and the failure. */
const rowsUntilFailure = async (book: X.TestBook, limits?: LimitOverrides) => {
  const wb = await Effect.runPromise(Xls.read(X.xls(book), { limits }))
  const seen: Array<Workbook.Row> = []
  const exit = await Effect.runPromise(Effect.exit(Stream.runForEach(Workbook.rows(wb, 0), (row) => Effect.sync(() => void seen.push(row)))))
  return { rows: seen.map((r) => r.index), failure: failureOf(exit) }
}
/** Offset of record `k` of a one-sheet book's sheet (its BOF is record -1). */
const recordOffset = (book: X.TestBook, k: number): number => {
  const stream = X.workbookStream(book)
  const sheet = book.sheets[0]!
  const before = X.workbookStream({ ...book, sheets: [{ ...sheet, records: sheet.records.slice(0, k) }] })
  // the shorter stream's EOF is where record k is
  expect(stream.length).toBeGreaterThan(before.length - 4)
  return before.length - 4
}
const one = (records: ReadonlyArray<Uint8Array>, globals: ReadonlyArray<Uint8Array> = [], biff: 5 | 8 = 8): X.TestBook => ({ biff, globals, sheets: [{ name: "S", records }] })

const values = (rows: ReadonlyArray<Workbook.Row>) => rows.map((r) => [r.index, r.hidden, r.cells.map((c) => [c.column, c._tag, Cell.value(c)])])
/** The cells of a one-sheet book, as [row, column, tag, value]. */
const cellsOf = async (records: ReadonlyArray<Uint8Array>, biff: 5 | 8 = 8) =>
  (await readOk(one(records, [], biff))).sheets[0]!.flatMap((r) => r.cells.map((c) => [r.index, c.column, c._tag, Cell.value(c)]))

const SHRFMLA = 0x04BC
const ARRAY = 0x0221
const TABLE = 0x0236
const CONTINUE = 0x003C
const STRING = 0x0207
/** A formula's companion record (content isn't read). */
const companion = (type: number) => X.record(type, new Array(14).fill(0))
/** FORMULA with the BIFF3/4 record number 0x0406: row, column, XF, result, flags, formula. */
const formula406 = (r: number, c: number, result: ReadonlyArray<number>) =>
  X.record(0x0406, [...X.u16(r), ...X.u16(c), ...X.u16(0), ...result, ...X.u16(0), ...X.u16(3), 0x1E, 1, 0])
const specialResult = (kind: number, value = 0) => [kind, 0, value, 0, 0, 0, 0xFF, 0xFF]
/** FORMULA with raw result bytes (e.g. a NaN with any payload). */
const formulaRaw = (r: number, c: number, result: ReadonlyArray<number>) =>
  X.record(0x0006, [...X.u16(r), ...X.u16(c), ...X.u16(0), ...result, ...X.u16(0), ...X.u32(0), ...X.u16(3), 0x1E, 1, 0])
const chars = (s: string) => Array.from(s, (c) => c.charCodeAt(0))
const wideChars = (s: string) => Array.from(s, (c) => c.charCodeAt(0)).flatMap(X.u16)

describe("xls sheet: formula string results", () => {
  it("a STRING after the formula's SHRFMLA, ARRAY or TABLE record is its result", async () => {
    for (const type of [SHRFMLA, ARRAY, TABLE]) {
      expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), companion(type), X.string("shared"), X.number(1, 0, 0, 2)]), type.toString(16)).toEqual([
        [0, 0, "Text", "shared"],
        [1, 0, "Number", 2]
      ])
    }
    // any other record in between ends the wait: the result is "", the STRING is ignored
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), X.number(0, 1, 0, 1), X.string("late")])).toEqual([
      [0, 0, "Text", ""],
      [0, 1, "Number", 1]
    ])
  })

  it("an embedded substream (BOF ... EOF) between FORMULA and STRING ends the wait, like any other record", async () => {
    // the STRING after the embedded chart isn't the formula's result
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), X.bof(0x0020), X.eof(), X.string("late"), X.number(1, 0, 0, 2)])).toEqual([
      [0, 0, "Text", ""],
      [1, 0, "Number", 2]
    ])
    // a STRING after it isn't read at all: one cut short is no failure
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), X.bof(0x0020), X.eof(), X.record(STRING, [...X.u16(9), 0, 0x41])])).toEqual([[0, 0, "Text", ""]])
    // nor is a STRING inside it, or the one after it
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), X.bof(0x0020), X.string("chart"), X.eof(), X.string("late"), X.number(1, 0, 0, 2)])).toEqual([
      [0, 0, "Text", ""],
      [1, 0, "Number", 2]
    ])
  })

  it("the wait for a formula's STRING, record by record: kept by SHRFMLA, ARRAY, TABLE and CONTINUE, taken by STRING, ended by any other record", async () => {
    const around = (records: ReadonlyArray<Uint8Array>) => cellsOf([X.number(0, 0, 0, 1), ...records, X.number(1, 0, 0, 2)])
    const filler = X.record(0x0225, X.u16(0))
    // no wait: a STRING, a companion record or a CONTINUE gives no cell
    for (const record of [X.string("x"), companion(SHRFMLA), companion(ARRAY), companion(TABLE), X.record(CONTINUE, [0, 0]), filler]) {
      expect(await around([record])).toEqual([[0, 0, "Number", 1], [1, 0, "Number", 2]])
    }
    // waiting: what comes after the FORMULA, and the formula's text
    const cases: ReadonlyArray<readonly [string, ReadonlyArray<Uint8Array>, string]> = [
      ["STRING", [X.string("s")], "s"],
      ["SHRFMLA, STRING", [companion(SHRFMLA), X.string("s")], "s"],
      ["ARRAY, STRING", [companion(ARRAY), X.string("s")], "s"],
      ["TABLE, STRING", [companion(TABLE), X.string("s")], "s"],
      ["CONTINUE, STRING", [X.record(CONTINUE, [0, 0]), X.string("s")], "s"],
      ["another record, STRING", [filler, X.string("s")], ""],
      ["another record", [filler], ""],
      ["an embedded substream, STRING", [X.bof(0x0020), X.eof(), X.string("s")], ""]
    ]
    for (const [name, records, text] of cases) {
      expect(await around([X.formula(0, 1, 0, { kind: 0 }), ...records]), name).toEqual([
        [0, 0, "Number", 1],
        [0, 1, "Text", text],
        [1, 0, "Number", 2]
      ])
    }
  })

  it("a FORMULA while another waits ends that wait, then waits itself; a wait open at the sheet's EOF gives \"\"", async () => {
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), X.formula(0, 1, 0, { kind: 0 }), X.string("s")])).toEqual([
      [0, 0, "Text", ""],
      [0, 1, "Text", "s"]
    ])
    expect(await cellsOf([X.number(0, 0, 0, 1), X.formula(0, 1, 0, { kind: 0 })])).toEqual([
      [0, 0, "Number", 1],
      [0, 1, "Text", ""]
    ])
  })

  it("a STRING continued by CONTINUE records, plain and switching width at the boundary", async () => {
    // "abcdef": 3 compressed characters, then a CONTINUE with option byte 0 and 3 more
    const plain = [X.record(STRING, [...X.u16(6), 0, ...chars("abc")]), X.record(CONTINUE, [0, ...chars("def")])]
    // "abcдеж": compressed, then a CONTINUE whose option byte switches to 16-bit
    const wide = [X.record(STRING, [...X.u16(6), 0, ...chars("abc")]), X.record(CONTINUE, [1, ...wideChars("деж")])]
    // 16-bit first, compressed after, over two CONTINUEs (an empty one in between continues nothing)
    const back = [
      X.record(STRING, [...X.u16(4), 1, ...wideChars("я")]),
      X.record(CONTINUE, []),
      X.record(CONTINUE, [0, ...chars("x")]),
      X.record(CONTINUE, [1, ...wideChars("ыz")])
    ]
    for (const [records, text] of [[plain, "abcdef"], [wide, "abcдеж"], [back, "яxыz"]] as const) {
      expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), ...records, X.number(1, 0, 0, 2)]), text).toEqual([
        [0, 0, "Text", text],
        [1, 0, "Number", 2]
      ])
    }
    // after an ARRAY, and a STRING cut short (fewer characters than its count) is MalformedFile
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), companion(ARRAY), ...wide])).toEqual([[0, 0, "Text", "abcдеж"]])
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), X.record(STRING, [...X.u16(6), 0, ...chars("abc")]), X.record(CONTINUE, [0, ...chars("d")])])))
      .toMatchObject({ _tag: "MalformedFile", reason: "truncated record 0x207" })
    // a CONTINUE of the STRING that runs past the end of the stream is that record's error, not the string's
    const cut = new Uint8Array([0x3C, 0x00, 0xFF, 0x00, 0x00, 0x62, 0x63])
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), X.record(STRING, [...X.u16(3), 0, ...chars("a")]), cut])))
      .toMatchObject({ _tag: "MalformedFile", reason: "record 0x3c runs past the end of the stream" })
  })

  it("a STRING's characters don't continue past the end of its sheet's substream", async () => {
    // sheet S ends with a STRING short of its count; the CONTINUE with the rest is where sheet T starts
    const short = X.record(STRING, [...X.u16(3), 0, ...chars("a")])
    const stream = X.workbookStream({ sheets: [{ name: "S", records: [X.formula(0, 0, 0, { kind: 0 }), short, X.record(CONTINUE, [0, ...chars("bc")])] }, { name: "T", records: [] }] })
    const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength)
    /** Offsets of the records from `at` on, up to the end of the stream. */
    const headers = (at: number): Array<number> => {
      const out: Array<number> = []
      // progress: each record moves `at` by 4 + its length
      for (; at + 4 <= stream.length; at += 4 + view.getUint16(at + 2, true)) out.push(at)
      return out
    }
    const bounds = headers(0).filter((at) => view.getUint16(at, true) === 0x0085)
    const sheetS = view.getUint32(bounds[0]! + 4, true)
    const string = headers(sheetS).find((at) => view.getUint16(at, true) === STRING)!
    const continued = string + 4 + view.getUint16(string + 2, true)
    expect(view.getUint16(continued, true)).toBe(CONTINUE)
    // point sheet T at the CONTINUE: sheet S's substream now ends right before it
    view.setUint32(bounds[1]! + 4, continued, true)
    expect(failureOf(await readExit(stream))).toMatchObject({ _tag: "MalformedFile", reason: "truncated record 0x207", offset: string })
  })

  it("a continued STRING over maxStringLength is LimitExceeded with its character count, however many bytes follow", async () => {
    const long = [X.record(STRING, [...X.u16(20), 0, ...chars("0123456789")]), X.record(CONTINUE, [0, ...chars("0123456789")])]
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), ...long]), { maxStringLength: 10 })).toMatchObject({ limit: "maxStringLength", max: 10, actual: 20 })
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), ...long])).toEqual([[0, 0, "Text", "01234567890123456789"]])
    // CONTINUEs far over any string of maxStringLength characters: the count is what is checked
    const filler = Array.from({ length: 40 }, () => X.record(CONTINUE, new Array(8000).fill(0)))
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), X.record(STRING, [...X.u16(15), 0, 0x41]), ...filler]), { maxStringLength: 10 }))
      .toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", max: 10, actual: 15 })
    // the count is checked before any character is read: characters cut short don't make it Truncated
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), X.record(STRING, [...X.u16(15), 0, 0x41, 0x42])]), { maxStringLength: 10 }))
      .toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", max: 10, actual: 15 })
    // a count within the limit reads the string: the CONTINUEs after its characters are passed over
    const short = [X.formula(0, 0, 0, { kind: 0 }), X.record(STRING, [...X.u16(1), 0, 0x41]), ...filler, X.number(1, 0, 0, 2)]
    expect(values((await readOk(one(short), { maxStringLength: 10 })).sheets[0]!)).toEqual([[0, false, [[0, "Text", "A"]]], [1, false, [[0, "Number", 2]]]])
  })

  it("BIFF5: STRING and LABEL are u16-count code-page strings; RSTRING (BIFF8) skips its runs", async () => {
    const cp1251 = X.record(0x0042, X.u16(1251))
    const book = one([
      X.formula(0, 0, 0, { kind: 0 }),
      X.record(STRING, X.string5([0xC0, 0xE1, 0xE2], 2)), // "Абв" in windows-1251
      X.formula(1, 0, 0, { kind: 0 }),
      companion(SHRFMLA),
      X.record(STRING, X.string5([0x41], 2))
    ], [cp1251], 5)
    const { sheets } = await readOk(book)
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Text", "Абв"]]], [1, false, [[0, "Text", "A"]]]])
    // RSTRING: XLUnicodeString, then a run count and runs
    const rstring = X.record(0x00D6, [...X.u16(0), ...X.u16(2), ...X.u16(0), ...X.xlString("rich"), ...X.u16(1), ...X.u16(0), ...X.u16(1)])
    expect(await cellsOf([rstring])).toEqual([[0, 2, "Text", "rich"]])
  })

  it("a FORMULA, SHRFMLA, ARRAY or TABLE continued by a CONTINUE record still takes the STRING after it", async () => {
    // the formula's own bytes split across a CONTINUE: the CONTINUE is part of it, not a record in between
    const data = Array.from(X.formula(0, 0, 0, { kind: 0 }).subarray(4))
    expect(await cellsOf([X.record(0x0006, data.slice(0, -1)), X.record(CONTINUE, data.slice(-1)), X.string("res")])).toEqual([[0, 0, "Text", "res"]])
    for (const type of [SHRFMLA, ARRAY, TABLE]) {
      const split = [X.record(type, new Array(10).fill(0)), X.record(CONTINUE, new Array(4).fill(0))]
      expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), ...split, X.string("shared"), X.number(1, 0, 0, 2)]), type.toString(16)).toEqual([
        [0, 0, "Text", "shared"],
        [1, 0, "Number", 2]
      ])
    }
  })

  it("FORMULA with the BIFF4 record number 0x0406 reads like FORMULA", async () => {
    expect(await cellsOf([
      formula406(0, 0, X.f64(3)),
      formula406(0, 1, specialResult(0)),
      X.string("text"),
      formula406(0, 2, specialResult(1, 1)),
      formula406(0, 3, specialResult(2, 0x07))
    ])).toEqual([
      [0, 0, "Number", 3],
      [0, 1, "Text", "text"],
      [0, 2, "Boolean", true],
      [0, 3, "Error", "#DIV/0!"]
    ])
  })
})

describe("xls sheet: numbers and errors", () => {
  it("NaN and infinities are #NUM! in NUMBER, RK, MULRK and FORMULA; -0 is 0", async () => {
    const nanPayload = [1, 0, 0, 0, 0, 0, 0xF0, 0x7F] // a signalling NaN, not 0xFFFF in its top bytes
    const rkInfinity = 0x7FF00000 // RK: the high 30 bits of +Infinity
    const cells = await cellsOf([
      X.number(0, 0, 0, NaN),
      X.number(0, 1, 0, -Infinity),
      X.rk(0, 2, 0, rkInfinity),
      X.mulrk(0, 3, [[0, rkInfinity], [0, X.rkInt(4)]]),
      X.formula(0, 5, 0, Infinity),
      formulaRaw(0, 6, nanPayload),
      X.formula(0, 7, 0, -0),
      X.number(0, 8, 0, -0)
    ])
    expect(cells).toEqual([
      [0, 0, "Error", "#NUM!"],
      [0, 1, "Error", "#NUM!"],
      [0, 2, "Error", "#NUM!"],
      [0, 3, "Error", "#NUM!"],
      [0, 4, "Number", 4],
      [0, 5, "Error", "#NUM!"],
      [0, 6, "Error", "#NUM!"],
      [0, 7, "Number", 0],
      [0, 8, "Number", 0]
    ])
    expect(Object.is(cells[7]![3], 0) && Object.is(cells[8]![3], 0)).toBe(true)
  })

  it("error codes in BOOLERR and FORMULA; an unknown code is #N/A in both", async () => {
    const codes: Array<[number, string]> = [
      [0x00, "#NULL!"],
      [0x07, "#DIV/0!"],
      [0x0F, "#VALUE!"],
      [0x17, "#REF!"],
      [0x1D, "#NAME?"],
      [0x24, "#NUM!"],
      [0x2A, "#N/A"],
      [0x2B, "#GETTING_DATA"],
      [0x99, "#N/A"],
      [0x01, "#N/A"]
    ]
    const records = codes.flatMap(([code], k) => [X.boolErr(k, 0, 0, code, true), X.formula(k, 1, 0, { kind: 2, value: code })])
    expect(await cellsOf(records)).toEqual(codes.flatMap(([, text], k) => [[k, 0, "Error", text], [k, 1, "Error", text]]))
    // BOOLERR booleans: any non-zero value is true
    expect(await cellsOf([X.boolErr(0, 0, 0, 0, false), X.boolErr(0, 1, 0, 2, false)])).toEqual([[0, 0, "Boolean", false], [0, 1, "Boolean", true]])
  })
})

describe("xls sheet: rows whose cells come in several places", () => {
  it("cells of a row before and after other rows' cells are merged into one row, in row order", async () => {
    const { sheets } = await readOk(one([
      X.row(0), X.row(1), X.row(2, true), X.row(3, true),
      X.number(1, 0, 0, 10),
      X.number(0, 1, 0, 1),
      X.number(1, 1, 0, 11),
      X.number(0, 0, 0, 0),
      X.number(4, 0, 0, 40),
      X.number(1, 2, 0, 12)
    ]))
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Number", 0], [1, "Number", 1]]],
      [1, false, [[0, "Number", 10], [1, "Number", 11], [2, "Number", 12]]],
      [2, true, []],
      [3, true, []],
      [4, false, [[0, "Number", 40]]]
    ])
    // a second pass over all rows (as some writers emit it), counted once per row against maxRows
    const pass = (c: number) => Array.from({ length: 150 }, (_, r) => X.number(r, c, 0, r))
    const twice = (await readOk(one([...pass(0), ...pass(1)]), { maxRows: 150 })).sheets[0]!
    expect(twice.map((r) => r.index)).toEqual(Array.from({ length: 150 }, (_, r) => r))
    expect(twice.every((r) => r.cells.length === 2 && r.cells[1]!.column === 1)).toBe(true)
  })

  it("a hidden ROW record after its row's cells still hides the row", async () => {
    const { sheets } = await readOk(one([X.number(0, 0, 0, 1), X.number(1, 0, 0, 2), X.row(0, true)]))
    expect(values(sheets[0]!)).toEqual([
      [0, true, [[0, "Number", 1]]],
      [1, false, [[0, "Number", 2]]]
    ])
  })

  it("ROW records after the cells are never out of place: a cap of 0 records out of place still reads them", async () => {
    // 20 rows of cells, then a ROW record for an empty row before them
    const empty = [...Array.from({ length: 20 }, (_, r) => X.number(r + 1, 0, 0, r)), X.row(0)]
    expect((await rowsDisplaced(one(empty), 0)).map((r) => [r.index, r.hidden, r.cells.length])).toEqual([
      [0, false, 0],
      ...Array.from({ length: 20 }, (_, r) => [r + 1, false, 1])
    ])
    // a hiding ROW record for the first row with cells, and a visible one repeated, both after all the cells
    const late = [X.row(0), ...Array.from({ length: 21 }, (_, r) => X.number(r, 0, 0, r)), X.row(0, true), X.row(1)]
    expect((await rowsDisplaced(one(late), 0)).map((r) => [r.index, r.hidden])).toEqual(Array.from({ length: 21 }, (_, r) => [r, r === 0]))
    // rows in reverse order, each row's cells together: none is out of place
    const reversed = Array.from({ length: 30 }, (_, r) => [X.number(29 - r, 0, 0, r), X.number(29 - r, 1, 0, r)]).flat()
    expect((await rowsDisplaced(one(reversed), 0)).map((r) => [r.index, r.cells.length])).toEqual(Array.from({ length: 30 }, (_, r) => [r, 2]))
  })

  it("rows declared only by ROW records come out in row order among the others, whatever the record order", async () => {
    const after = (await readOk(one([X.number(5, 0, 0, 1), X.row(2, true)]))).sheets[0]!
    expect(values(after)).toEqual([[2, true, []], [5, false, [[0, "Number", 1]]]])
    const descending = (await readOk(one([X.row(3), X.row(1, true)]))).sheets[0]!
    expect(values(descending)).toEqual([[1, true, []], [3, false, []]])
    const mixed = (await readOk(one([X.row(4), X.number(2, 0, 0, 1), X.row(0, true), X.number(3, 0, 0, 2), X.row(1)]))).sheets[0]!
    expect(mixed.map((r) => [r.index, r.hidden])).toEqual([[0, true], [1, false], [2, false], [3, false], [4, false]])
  })

  it("a row waits for its formula's STRING even when the row's cells come back later", async () => {
    expect(await cellsOf([
      X.number(1, 0, 0, 1),
      X.formula(0, 0, 0, { kind: 0 }),
      companion(SHRFMLA),
      X.string("s"),
      X.number(1, 1, 0, 2)
    ])).toEqual([
      [0, 0, "Text", "s"],
      [1, 0, "Number", 1],
      [1, 1, "Number", 2]
    ])
  })

  it("a FORMULA, its STRING, records between them and an embedded substream are never out of place; a cell record back in its row is", async () => {
    // row 0's last BLANK comes after row 1: the one record out of place
    const records = [X.blank(0, 0, 0), X.blank(1, 0, 0), X.blank(1, 1, 0), X.formula(1, 2, 0, { kind: 0 }), companion(SHRFMLA), X.string("s"), X.bof(0x0020), X.number(0, 5, 0, 1), X.eof(), X.blank(0, 1, 0)]
    const book = one(records)
    expect(failureOf(await rowsDisplacing(book, 0))).toMatchObject({
      _tag: "MalformedFile",
      offset: recordOffset(book, 9),
      reason: "rows split too far apart: over 0 cell records out of their row's place"
    })
    expect(values(await rowsDisplaced(book, 1))).toEqual([
      [0, false, [[0, "Empty", undefined], [1, "Empty", undefined]]],
      [1, false, [[0, "Empty", undefined], [1, "Empty", undefined], [2, "Text", "s"]]]
    ])
  })

  it("a row whose last cell is a formula with a string result is emitted only after its STRING", async () => {
    const { sheets } = await readOk(one([X.number(0, 0, 0, 1), X.formula(0, 1, 0, { kind: 0 }), X.string("s"), X.number(1, 0, 0, 2)]))
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Number", 1], [1, "Text", "s"]]],
      [1, false, [[0, "Number", 2]]]
    ])
  })

  it("a row split far apart reads, its cells merged in stream order; past the cap on records out of place the sheet is MalformedFile", async () => {
    // a BLANK in row 0, rows 1-20 of 100 cells each, then row 0's NUMBER in the same column and a BLANK:
    // two records out of place, the NUMBER the last cell of column 0
    const records = [X.blank(0, 0, 0), ...Array.from({ length: 20 }, (_, r) => X.mulrk(r + 1, 0, Array.from({ length: 100 }, (_, c) => [0, X.rkInt(c)] as const))), X.number(0, 0, 0, 7), X.blank(0, 1, 0)]
    const rows = await rowsDisplaced(one(records), 2)
    expect(rows.map((r) => [r.index, r.cells.length])).toEqual([[0, 2], ...Array.from({ length: 20 }, (_, r) => [r + 1, 100])])
    expect(values(rows.slice(0, 1))).toEqual([[0, false, [[0, "Number", 7], [1, "Empty", undefined]]]])
    // records out of place in the same column: the later one wins
    const twice = [X.blank(0, 0, 0), X.number(1, 0, 0, 0), X.number(0, 1, 0, 1), X.number(1, 1, 0, 0), X.number(0, 1, 0, 2)]
    expect(values(await rowsDisplaced(one(twice), 3))).toEqual([
      [0, false, [[0, "Empty", undefined], [1, "Number", 2]]],
      [1, false, [[0, "Number", 0], [1, "Number", 0]]]
    ])
    // a cap of 1: the second record out of place fails, where it is
    const book = one(records)
    expect(failureOf(await rowsDisplacing(book, 1))).toMatchObject({
      _tag: "MalformedFile",
      offset: recordOffset(book, 22),
      reason: "rows split too far apart: over 1 cell records out of their row's place"
    })
    // the default cap
    expect(DISPLACED_RECORDS).toBe(1_048_576)
  })

  it("a row split around more cells than the old cap on held cells (1 Mi) reads: what is out of place costs an offset, not its cells", async () => {
    // row 0's first BLANK, rows 1-260 of 4 096 empty cells each (1 064 960 cells waiting for row 0), then row 0's last BLANK
    const xfs = new Array<number>(4096).fill(0)
    const records = [X.blank(0, 0, 0), ...Array.from({ length: 260 }, (_, r) => X.mulBlank(r + 1, 0, xfs)), X.blank(0, 1, 0)]
    const { sheets } = await readOk(one(records))
    expect(sheets[0]!.map((r) => r.cells.length)).toEqual([2, ...new Array(260).fill(4096)])
  }, 60_000)

  it("past the default cap on records out of place (1 Mi) the sheet is MalformedFile", async () => {
    // rows 0 and 1 alternating: after the first two, every BLANK is back in a row another row's cells ended
    const size = X.blank(0, 0, 0).length
    const count = DISPLACED_RECORDS + 3
    const blob = new Uint8Array(count * size)
    // progress: k up to count
    for (let k = 0; k < count; k++) blob.set(X.blank(k % 2, 0, 0), k * size)
    const book = one([blob])
    const failure = failureOf(await rowsDisplacing(book, DISPLACED_RECORDS))
    expect(failure).toMatchObject({ _tag: "MalformedFile", reason: `rows split too far apart: over ${DISPLACED_RECORDS} cell records out of their row's place` })
    // the record after the cap's last: 2 in place, then DISPLACED_RECORDS out of place
    expect((failure as { offset: number }).offset).toBe(recordOffset(book, 0) + (DISPLACED_RECORDS + 2) * size)
    // one record fewer reads
    const fewer = one([blob.subarray(0, (count - 1) * size)])
    expect((await rowsDisplaced(fewer, DISPLACED_RECORDS)).map((r) => [r.index, r.cells.length])).toEqual([[0, 1], [1, 1]])
  }, 60_000)

  it("a duplicate column in a row keeps the last cell", async () => {
    expect(await cellsOf([X.number(0, 3, 0, 1), X.number(0, 1, 0, 2), X.number(0, 3, 0, 3)])).toEqual([
      [0, 1, "Number", 2],
      [0, 3, "Number", 3]
    ])
    expect(await cellsOf([X.number(0, 0, 0, 1), X.number(0, 0, 0, 2)])).toEqual([[0, 0, "Number", 2]])
  })

  it("rows in order still stream: after the scan of the record headers, the first row comes out before the rest of the sheet's records are read", async () => {
    // row 0, then more than a batch of records, then a record past the end of the stream
    const filler = Array.from({ length: 2 * BATCH_RECORDS }, (_, k) => X.number(1 + k, 0, 0, k))
    const stream = X.workbookStream(one([X.number(0, 0, 0, 7), ...filler]))
    const cut = stream.subarray(0, stream.length - 10)
    const wb = await Effect.runPromise(Xls.read(writeCfb([{ name: "Workbook", data: cut }]).bytes))
    const first = await Effect.runPromise(Stream.runCollect(Stream.take(Workbook.rows(wb, 0), 1)))
    expect(Array.from(first).map((r) => r.index)).toEqual([0])
    expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)))))).toMatchObject({ _tag: "MalformedFile" })
  })
})

describe("xls sheet: emitting rows", () => {
  it("a row goes out only if it and every row before it were complete before the first failure in stream order", async () => {
    // rows 0 and 1 complete; row 2's last cell after a truncated NUMBER (in row 9); row 3 complete before it
    const cut = X.record(0x0203, [...X.u16(9), ...X.u16(0)])
    const records = [X.number(0, 0, 0, 1), X.number(1, 0, 0, 1), X.number(2, 0, 0, 1), X.number(3, 0, 0, 1), cut, X.number(2, 1, 0, 1)]
    const book = one(records)
    expect(await rowsUntilFailure(book)).toEqual({ rows: [0, 1], failure: expect.objectContaining({ reason: "truncated record 0x203", offset: recordOffset(book, 4) }) })
    // a ROW record cut before its flags fails where it is
    expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), X.record(0x0208, [...X.u16(3), 0, 0, 0, 0]), X.number(1, 0, 0, 1)])))
      .toEqual({ rows: [0], failure: expect.objectContaining({ reason: "truncated record 0x208" }) })
    // ROW-only rows are complete from the start, wherever their ROW record is
    expect(await rowsUntilFailure(one([X.row(0), X.row(1), cut, X.row(4)]))).toEqual({ rows: [0, 1, 4], failure: expect.objectContaining({ reason: "truncated record 0x203" }) })
    // a row whose FORMULA waits for its STRING until the failing record (which ends the wait) isn't complete
    const waiting = one([X.number(0, 0, 0, 1), X.formula(1, 0, 0, { kind: 0 }), companion(SHRFMLA), X.record(0x0201, [0])])
    expect(await rowsUntilFailure(waiting)).toEqual({ rows: [0], failure: expect.objectContaining({ reason: "truncated record 0x201" }) })
    const pending = one([X.number(0, 0, 0, 1), X.formula(1, 0, 0, { kind: 0 }), companion(SHRFMLA), X.record(STRING, [...X.u16(9), 0, 0x41])])
    expect(await rowsUntilFailure(pending)).toEqual({ rows: [0], failure: expect.objectContaining({ reason: "truncated record 0x207" }) })
    // still waiting when the records stop: neither is that row
    const stopped = X.workbookStream(one([X.number(0, 0, 0, 1), X.formula(1, 0, 0, { kind: 0 }), companion(SHRFMLA), X.number(2, 0, 0, 1)]))
    const wb = await Effect.runPromise(Xls.read(writeCfb([{ name: "Workbook", data: stopped.subarray(0, stopped.length - 10) }]).bytes))
    const seen: Array<number> = []
    const exit = await Effect.runPromise(Effect.exit(Stream.runForEach(Workbook.rows(wb, 0), (row) => Effect.sync(() => void seen.push(row.index)))))
    expect([seen, failureOf(exit)]).toEqual([[0], expect.objectContaining({ reason: "record 0x203 runs past the end of the stream" })])
    // the maxRows failure of a row comes once the rows before it are out, if no failure in stream order came first
    expect(await rowsUntilFailure(one([X.number(5, 0, 0, 1), X.number(0, 0, 0, 1), X.number(1, 0, 0, 1)]), { maxRows: 2 }))
      .toEqual({ rows: [0, 1], failure: expect.objectContaining({ limit: "maxRows", actual: 6 }) })
    expect(await rowsUntilFailure(one([X.number(5, 0, 0, 1), X.number(0, 0, 0, 1), cut, X.number(1, 0, 0, 1)]), { maxRows: 2 }))
      .toEqual({ rows: [0], failure: expect.objectContaining({ reason: "truncated record 0x203" }) })
  })

  it("the failure is the first in stream order, whatever row its record is in; cells are counted in stream order", async () => {
    // row 5's LABEL over maxStringLength comes before row 0's truncated record
    const long = X.label(5, 0, 0, "a string too long")
    const cut = X.record(0x0203, [...X.u16(0), ...X.u16(1)])
    expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), long, cut]), { maxStringLength: 10 }))
      .toEqual({ rows: [], failure: expect.objectContaining({ limit: "maxStringLength", actual: 17 }) })
    // so is row 5's missing shared string, and row 5's formula result cut short
    expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), X.labelSst(5, 0, 0, 3), cut])))
      .toEqual({ rows: [], failure: expect.objectContaining({ reason: "shared string 3 doesn't exist" }) })
    const shortString = X.record(STRING, [...X.u16(9), 0, 0x41])
    expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), X.formula(5, 0, 0, { kind: 0 }), shortString, cut])))
      .toEqual({ rows: [], failure: expect.objectContaining({ reason: "truncated record 0x207" }) })
    // row 1's three cells pass maxCellsPerSheet before row 0's cell in column 10 passes maxColumns
    const cells = [X.number(1, 0, 0, 1), X.number(1, 1, 0, 1), X.number(1, 2, 0, 1), X.number(0, 10, 0, 1)]
    expect(await rowsUntilFailure(one(cells), { maxCellsPerSheet: 2, maxColumns: 5 }))
      .toEqual({ rows: [], failure: expect.objectContaining({ limit: "maxCellsPerSheet", actual: 3 }) })
    expect(await rowsUntilFailure(one(cells.slice(1).concat(cells.slice(0, 1))), { maxCellsPerSheet: 3, maxColumns: 5 }))
      .toEqual({ rows: [], failure: expect.objectContaining({ limit: "maxColumns", actual: 11 }) })
    // within a MULRK, cell by cell: a column at maxColumns before the count, the count before a later column
    const mul = X.mulrk(0, 3, [[0, 1], [0, 2], [0, 3]])
    expect(await errorOf(one([mul]), { maxColumns: 6, maxCellsPerSheet: 2 })).toMatchObject({ limit: "maxCellsPerSheet", actual: 3 })
    expect(await errorOf(one([mul]), { maxColumns: 5, maxCellsPerSheet: 3 })).toMatchObject({ limit: "maxColumns", actual: 6 })
    expect(await errorOf(one([mul]), { maxColumns: 4, maxCellsPerSheet: 1 })).toMatchObject({ limit: "maxColumns", actual: 5 })
    // a formula's "" result counts where its wait ends
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), X.number(1, 0, 0, 1)]), { maxCellsPerSheet: 1 })).toMatchObject({ limit: "maxCellsPerSheet", actual: 2 })
  })

  it("a FORMULA out of place takes its STRING, right after it or after a SHRFMLA, ARRAY, TABLE or CONTINUE", async () => {
    for (const between of [[], [companion(SHRFMLA)], [companion(TABLE), X.record(CONTINUE, [0, 0])], [X.record(CONTINUE, [1])]]) {
      expect(await cellsOf([X.number(0, 0, 0, 1), X.number(1, 0, 0, 2), X.formula(0, 1, 0, { kind: 0 }), ...between, X.string("late"), X.number(2, 0, 0, 3)]), String(between.length)).toEqual([
        [0, 0, "Number", 1],
        [0, 1, "Text", "late"],
        [1, 0, "Number", 2],
        [2, 0, "Number", 3]
      ])
    }
    // in place, and without a STRING after its SHRFMLA: ""
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), companion(ARRAY), X.formula(0, 1, 0, { kind: 0 }), companion(SHRFMLA), X.string("b")])).toEqual([
      [0, 0, "Text", ""],
      [0, 1, "Text", "b"]
    ])
  })

  it("ROW-only rows come out in chunks of at most BATCH_RECORDS rows", async () => {
    const wb = await Effect.runPromise(Xls.read(X.xls(one(Array.from({ length: 10 * BATCH_RECORDS }, (_, r) => X.row(r))))))
    const chunks = Array.from(await Effect.runPromise(Stream.runCollect(Stream.chunks(Workbook.rows(wb, 0)))))
    expect(chunks.reduce((n, c) => n + c.length, 0)).toBe(10 * BATCH_RECORDS)
    expect(Math.max(...chunks.map((c) => c.length))).toBe(BATCH_RECORDS)
  })

  it("repeated ROW records give one row each, hidden if any of them hides it", async () => {
    const flood = Array.from({ length: 10_000 }, (_, k) => X.row(k % 2, k === 7))
    expect(values((await readOk(one(flood))).sheets[0]!)).toEqual([[0, false, []], [1, true, []]])
  })

  it("a row waiting for its formula's STRING doesn't stop the reader: the read ends at once (pending STRING isn't behind)", async () => {
    // if behind() took a row waiting for its STRING as due, the reader
    // would stop reading and loop (yielding) forever: the timeout catches
    // it in a second instead of the test's own
    const book = one([X.number(0, 0, 0, 1), X.formula(0, 1, 0, { kind: 0 }), companion(SHRFMLA), X.string("s"), X.number(1, 0, 0, 2)])
    const wb = await Effect.runPromise(Xls.read(X.xls(book)))
    const exit = await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)).pipe(Effect.timeout("1 second"))))
    if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
    expect(values(Array.from(exit.value))).toEqual([
      [0, false, [[0, "Number", 1], [1, "Text", "s"]]],
      [1, false, [[0, "Number", 2]]]
    ])
  })
})

describe("xls sheet: the index pass checks every record as the rows read it", () => {
  /** A NUMBER in row 0 cut short (4 bytes): the failure if nothing before it fails. */
  const cutNumber = X.record(0x0203, [...X.u16(0), ...X.u16(1)])
  /** The record `record` with its data cut to `length` bytes. */
  const cutTo = (record: Uint8Array, length: number) => X.record(record[0]! | (record[1]! << 8), Array.from(record.subarray(4, 4 + length)))

  it("each cell record is checked against its own minimum length: one byte short fails where it is, before a later failure in an earlier row", async () => {
    const cases: ReadonlyArray<readonly [Uint8Array, number]> = [
      [X.number(5, 0, 0, 1), 14],
      [X.rk(5, 0, 0, X.rkInt(1)), 10],
      [X.boolErr(5, 0, 0, 1, false), 8],
      [X.labelSst(5, 0, 0, 0), 10],
      [X.formula(5, 0, 0, 1), 14],
      [X.blank(5, 0, 0), 6],
      [X.label(5, 0, 0, "x"), 6],
      [X.mulBlank(5, 0, [0, 0]), 6]
    ]
    for (const [record, min] of cases) {
      const type = record[0]! | (record[1]! << 8)
      // the later failure: row 0's missing shared string
      expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), cutTo(record, min - 1), X.labelSst(0, 1, 0, 99)]), {}), type.toString(16))
        .toEqual({ rows: [], failure: expect.objectContaining({ reason: `truncated record 0x${type.toString(16)}`, offset: expect.any(Number) }) })
    }
  })

  it("a MULRK or MULBLANK counts every one of its cells; a formula's STRING result counts its cell", async () => {
    expect(await rowsUntilFailure(one([X.mulBlank(0, 0, [0, 0, 0]), X.number(1, 0, 0, 1)]), { maxCellsPerSheet: 3 }))
      .toEqual({ rows: [0], failure: expect.objectContaining({ limit: "maxCellsPerSheet", actual: 4 }) })
    expect(await rowsUntilFailure(one([X.mulrk(0, 0, [[0, 1], [0, 2]]), X.number(1, 0, 0, 1)]), { maxCellsPerSheet: 2 }))
      .toEqual({ rows: [0], failure: expect.objectContaining({ limit: "maxCellsPerSheet", actual: 3 }) })
    expect(await rowsUntilFailure(one([X.formula(0, 0, 0, { kind: 0 }), X.string("a"), X.number(1, 0, 0, 1)]), { maxCellsPerSheet: 1 }))
      .toEqual({ rows: [0], failure: expect.objectContaining({ limit: "maxCellsPerSheet", actual: 2 }) })
  })

  it("only a FORMULA with a string result waits for a STRING: other special results leave it unread", async () => {
    for (const kind of [1, 2, 3] as const) {
      const exit = await rowsUntilFailure(one([X.formula(0, 0, 0, { kind, value: 1 }), X.string("much too long"), X.number(1, 0, 0, 1)]), { maxStringLength: 4 })
      expect(exit, String(kind)).toEqual({ rows: [0, 1], failure: undefined })
    }
    // a wait ended by another record after a SHRFMLA: "" (in a row other than the first)
    expect(await cellsOf([X.number(0, 0, 0, 1), X.formula(3, 0, 0, { kind: 0 }), companion(SHRFMLA), X.number(5, 0, 0, 1), X.string("late")])).toEqual([
      [0, 0, "Number", 1],
      [3, 0, "Text", ""],
      [5, 0, "Number", 1]
    ])
  })

  it("a formula's STRING is checked without decoding it, in stream order: BIFF8 rich runs and phonetic block, BIFF5 length", async () => {
    // cch 1, fRichSt with 4 runs, one character, no run data
    const rich = X.record(STRING, [...X.u16(1), 0x08, ...X.u16(4), 0x41])
    // cch 1, fExtSt with a 6-byte phonetic block, one character, 2 of its bytes
    const ext = X.record(STRING, [...X.u16(1), 0x04, ...X.u32(6), 0x41, 0, 0])
    for (const cut of [rich, ext]) {
      expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), X.formula(5, 0, 0, { kind: 0 }), cut, cutNumber])))
        .toEqual({ rows: [], failure: expect.objectContaining({ reason: "truncated record 0x207" }) })
    }
    const long5 = X.record(STRING, X.string5(chars("abcdefghi"), 2))
    expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), X.formula(5, 0, 0, { kind: 0 }), long5, cutNumber], [], 5), { maxStringLength: 4 }))
      .toEqual({ rows: [], failure: expect.objectContaining({ limit: "maxStringLength", actual: 9 }) })
    // a BIFF5 LABEL or STRING of exactly maxStringLength reads
    const book = one([X.label5(0, 0, 0, chars("abcd")), X.formula(1, 0, 0, { kind: 0 }), X.record(STRING, X.string5(chars("wxyz"), 2))], [], 5)
    expect(values((await readOk(book, { maxStringLength: 4 })).sheets[0]!)).toEqual([[0, false, [[0, "Text", "abcd"]]], [1, false, [[0, "Text", "wxyz"]]]])
  })

  it("a ROW cut at 13 bytes (before its flags' high byte) still hides its row, which goes out before the failure", async () => {
    expect(await rowsUntilFailure(one([X.number(0, 0, 0, 1), cutTo(X.row(0, true), 13)])))
      .toEqual({ rows: [0], failure: expect.objectContaining({ reason: "truncated record 0x208" }) })
    const wb = await Effect.runPromise(Xls.read(X.xls(one([X.number(0, 0, 0, 1), cutTo(X.row(0, true), 13)]))))
    const seen: Array<boolean> = []
    await Effect.runPromise(Effect.exit(Stream.runForEach(Workbook.rows(wb, 0), (row) => Effect.sync(() => void seen.push(row.hidden)))))
    expect(seen).toEqual([true])
  })

  it("a ROW record between a row's cells doesn't put them out of place", async () => {
    expect((await rowsDisplaced(one([X.number(0, 0, 0, 1), X.row(5), X.number(0, 1, 0, 2)]), 0)).map((r) => [r.index, r.cells.length])).toEqual([[0, 2], [5, 0]])
  })
})

describe("xls sheet: the rows' work", () => {
  /** Drains sheet 0's rows: the most steps between two event-loop turns, and the total. */
  const drain = async (book: X.TestBook, limits?: LimitOverrides) => {
    const wb = await Effect.runPromise(Xls.read(X.xls(book), { limits }))
    const steps = counter("steps")
    const run = await countYields(() => Effect.runPromise(metered(steps, Stream.runDrain(Workbook.rows(wb, 0)))), steps.total)
    return { maxPerTurn: run.maxPerTurn, total: steps.total() }
  }

  it("a row of single-cell records repeating its columns is compacted as it is read: no long sort at its end", async () => {
    // 200 000 NUMBERs over columns 7..0 in one row: compacted past 16 cells each time
    const size = X.number(0, 0, 0, 0).length
    const blob = new Uint8Array(200_000 * size)
    // progress: k up to 200 000
    for (let k = 0; k < 200_000; k++) blob.set(X.number(0, 7 - (k % 8), 0, k), k * size)
    const { maxPerTurn } = await drain(one([blob]), { maxColumns: 8 })
    // a turn: a budget of BATCH_CELLS (a record costs 16), a row's slack, and a few compactions; never the 200 000 cells at once
    expect(maxPerTurn).toBeLessThan(20_000)
  }, 60_000)

  it("records inside a row's range that aren't its cells cost their share of a turn", async () => {
    // row 0's two cells around 100 000 records of another kind: one range, read again by the rows
    const filler = X.record(0x0225, X.u16(0))
    const blob = new Uint8Array(100_000 * filler.length)
    // progress: k up to 100 000
    for (let k = 0; k < 100_000; k++) blob.set(filler, k * filler.length)
    const { maxPerTurn } = await drain(one([X.number(0, 0, 0, 1), blob, X.number(0, 1, 0, 2)]))
    expect(maxPerTurn).toBeLessThan(20_000)
  }, 60_000)

  it("a FORMULA whose STRING comes after a long run of CONTINUEs: the run is read by the index pass in bounded turns, never again", async () => {
    const run = new Uint8Array(400_000 * 4)
    // progress: k up to 400 000
    for (let k = 0; k < 400_000; k++) run.set(X.record(CONTINUE, []), k * 4)
    const book = one([X.formula(0, 0, 0, { kind: 0 }), run, X.string("x"), X.number(1, 0, 0, 1)])
    const { maxPerTurn, total } = await drain(book)
    // a turn holds about BATCH_RECORDS records: the index pass's 400 004 records over about 100 turns
    expect(maxPerTurn).toBeLessThan(2 * BATCH_RECORDS)
    // the index pass reads 400 004 records (the EOF too); the rows read 2 records, add 2 cells and emit 2 rows:
    // the run isn't scanned again
    expect(total).toBe(400_004 + 2 + 2 + 2)
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), run.subarray(0, 4000), X.string("x")])).toEqual([[0, 0, "Text", "x"]])
  }, 60_000)

  it("a row's FORMULAs take their noted STRINGs in stream order, out-of-place ones too", async () => {
    expect(await cellsOf([
      X.formula(0, 0, 0, { kind: 0 }), companion(SHRFMLA), X.string("a"),
      X.formula(1, 0, 0, { kind: 0 }), companion(TABLE), X.string("r1"),
      X.formula(0, 1, 0, { kind: 0 }), companion(ARRAY), X.record(CONTINUE, [0]), X.string("b"),
      X.formula(0, 2, 0, { kind: 0 }), X.record(CONTINUE, [1]), X.number(2, 0, 0, 1),
      X.formula(0, 3, 0, { kind: 0 }), X.string("d")
    ])).toEqual([
      [0, 0, "Text", "a"],
      [0, 1, "Text", "b"],
      [0, 2, "Text", ""],
      [0, 3, "Text", "d"],
      [1, 0, "Text", "r1"],
      [2, 0, "Number", 1]
    ])
    // rows first seen in reverse order: each row starts at its own notes
    expect(await cellsOf([X.formula(1, 0, 0, { kind: 0 }), companion(SHRFMLA), X.string("r1"), X.formula(0, 0, 0, { kind: 0 }), companion(SHRFMLA), X.string("r0")])).toEqual([
      [0, 0, "Text", "r0"],
      [1, 0, "Text", "r1"]
    ])
  })

  it("many noted STRINGs, over rows alternating, are taken over several turns", async () => {
    // 200 000 FORMULAs, each with a CONTINUE before its STRING, over 2 rows
    const pair = [...X.formula(0, 0, 0, { kind: 0 }), ...X.record(CONTINUE, []), ...X.string("s")]
    const blob = new Uint8Array(200_000 * pair.length)
    // progress: k up to 200 000
    for (let k = 0; k < 200_000; k++) {
      blob.set(pair, k * pair.length)
      blob[k * pair.length + 4] = k % 2
      blob[k * pair.length + 6] = k % 200
    }
    const { maxPerTurn } = await drain(one([blob]))
    // a turn: about BATCH_RECORDS records, read or decoded
    expect(maxPerTurn).toBeLessThan(100_000)
  }, 60_000)

  it("a row's records out of place are read over several turns", async () => {
    // rows 0 and 1 alternating: 400 000 BLANKs, all but two out of place
    const size = X.blank(0, 0, 0).length
    const blob = new Uint8Array(400_000 * size)
    // progress: k up to 400 000
    for (let k = 0; k < 400_000; k++) blob.set(X.blank(k % 2, 0, 0), k * size)
    const { maxPerTurn } = await drain(one([blob]))
    // a turn: about 3 800 records of 2 steps each, and a compaction of the row (32 769 cells); not 400 000 records at once
    expect(maxPerTurn).toBeLessThan(100_000)
  }, 60_000)

  it("a turn finishes the wide row it is in: each chunk's turn did exactly its rows' work", async () => {
    // 64 rows of 16 384 cells (4 MULBLANKs each): a row costs about a quarter turn, so a turn's budget runs out inside a row
    const xfs = new Array<number>(4096).fill(0)
    const wb = await Effect.runPromise(Xls.read(X.xls(one(Array.from({ length: 64 }, (_, r) => [0, 1, 2, 3].map((k) => X.mulBlank(r, 4096 * k, xfs))).flat()))))
    // the work meter is told each turn's steps as its pull ends, right before the turn's chunk goes out
    const events: Array<readonly ["steps" | "rows", number]> = []
    const meter: Meter = { add: (unit, count) => void (unit === "steps" && events.push(["steps", count])) }
    await Effect.runPromise(metered(meter, Stream.runForEachArray(Workbook.rows(wb, 0), (chunk) => Effect.sync(() => void events.push(["rows", chunk.length])))))
    const turns: Array<readonly [number, number]> = []
    // progress: k up to events.length
    for (let k = 0; k + 1 < events.length; k++) {
      if (events[k]![0] === "steps" && events[k + 1]![0] === "rows") turns.push([events[k + 1]![1], events[k]![1]])
    }
    expect(turns.reduce((n, [rows]) => n + rows, 0)).toBe(64)
    // every turn after the first (which also did the index pass): 4 records, 16 384 cells and the row, per row
    expect(turns.slice(1).map(([rows, steps]) => steps - rows * (4 + 16_384 + 1))).toEqual(turns.slice(1).map(() => 0))
    expect(turns.length).toBeGreaterThan(10)
  }, 60_000)

  it("records out of place and noted STRINGs cost a step per record, cell and row, nothing more", async () => {
    // rows 0 and 1 alternating, 1 000 BLANKs over columns k % 7 (998 out of place): the index pass reads 1 001
    // records (the EOF too); the rows read the 1 000 records and add their cells (2 000), sort each row's 500
    // cells (out of column order) and emit 2 rows
    const alternating = Array.from({ length: 1000 }, (_, k) => X.blank(k % 2, k % 7, 0))
    expect((await drain(one(alternating))).total).toBe(1001 + 2000 + 1000 + 2)
    // two FORMULAs of row 0 with their STRINGs after a SHRFMLA (two notes): the index pass reads 8 records; row 0's
    // range runs from its first FORMULA to its second (4 records, 2 cells); row 1: a record and a cell; 2 rows
    const notes = [X.formula(0, 0, 0, { kind: 0 }), companion(SHRFMLA), X.string("a"), X.formula(0, 1, 0, { kind: 0 }), companion(SHRFMLA), X.string("b"), X.number(1, 0, 0, 1)]
    expect((await drain(one(notes))).total).toBe(8 + 4 + 2 + 1 + 1 + 2)
    // the same with row 1's FORMULA between them: row 0's second FORMULA is out of place (a record and a cell)
    const between = [X.formula(0, 0, 0, { kind: 0 }), companion(SHRFMLA), X.string("a"), X.formula(1, 0, 0, { kind: 0 }), companion(SHRFMLA), X.string("r"), X.formula(0, 1, 0, { kind: 0 }), companion(SHRFMLA), X.string("b")]
    expect((await drain(one(between))).total).toBe(10 + (1 + 1) + (1 + 1) + (1 + 1) + 2)
    expect(await cellsOf(between)).toEqual([[0, 0, "Text", "a"], [0, 1, "Text", "b"], [1, 0, "Text", "r"]])
  })

  it("sorting a row out of column order counts its cells", async () => {
    // 1 000 NUMBERs in columns 999..0, one row: the index pass reads 1 001 records (the EOF too);
    // the rows read 1 000 records and add 1 000 cells, then sort them and emit the row
    const records = Array.from({ length: 1000 }, (_, k) => X.number(0, 999 - k, 0, k))
    expect((await drain(one(records))).total).toBe(1001 + 2000 + 1000 + 1)
    // in column order: no sort
    const sorted = Array.from({ length: 1000 }, (_, k) => X.number(0, k, 0, k))
    expect((await drain(one(sorted))).total).toBe(1001 + 2000 + 1)
  })
})

describe("xls sheet: repeated columns", () => {
  it("a row's repeated columns are compacted, the last cell of a column kept", async () => {
    // read end to end: the same rule over many MULBLANK and NUMBER records on the same columns
    const records = Array.from({ length: 50 }, (_, k) => [X.mulBlank(0, 0, new Array(8).fill(0)), X.number(0, 7 - (k % 8), 0, k)]).flat()
    const { sheets } = await readOk(one(records), { maxColumns: 8 })
    // the last writes: the last MULBLANK everywhere, then NUMBER 49 in column 7 - 49 % 8 = 6
    expect(sheets[0]![0]!.cells.map((c) => [c.column, c._tag, Cell.value(c) ?? null])).toEqual(
      Array.from({ length: 8 }, (_, c) => (c === 6 ? [6, "Number", 49] : [c, "Empty", null]))
    )
  })
})

describe("xls sheet: sheetDetails", () => {
  it("reversed merges are normalized; COLINFO with last < first and an empty DIMENSIONS are ignored", async () => {
    const { details } = await readOk(one([
      X.dimensions(3, 3, 0, 4), // no rows: r2 is exclusive
      X.colInfo(5, 2, true),
      X.colInfo(7, 7, true),
      X.colInfo(8, 9, false),
      X.mergeCells([[4, 2, 6, 1]])
    ]))
    expect(details[0]).toEqual({ hiddenColumns: [{ first: 7, last: 7 }], merges: [{ firstRow: 2, firstColumn: 1, lastRow: 4, lastColumn: 6 }] })
    // empty column range too; a one-cell dimension is kept (BIFF5 layout)
    expect((await readOk(one([X.dimensions(0, 5, 2, 2)]))).details[0]!.dimension).toBeUndefined()
    expect((await readOk(one([X.dimensions(1, 2, 2, 3, 5)], [], 5))).details[0]!.dimension).toEqual({ firstRow: 1, firstColumn: 2, lastRow: 1, lastColumn: 2 })
  })

  it("a COLINFO starting at or past maxColumns is ignored; one crossing it is clamped", async () => {
    const { details } = await readOk(one([X.colInfo(256, 300, true), X.colInfo(400, 400, true), X.colInfo(250, 260, true)]), { maxColumns: 256 })
    expect(details[0]!.hiddenColumns).toEqual([{ first: 250, last: 255 }])
  })

  it("a MERGECELLS shorter than its count is MalformedFile; sheetDetails yields over a long sheet", async () => {
    const short = X.record(0x00E5, [...X.u16(2), ...X.u16(0), ...X.u16(1), ...X.u16(0), ...X.u16(1)])
    expect(await errorOf(one([short]))).toMatchObject({ _tag: "MalformedFile", reason: "truncated MERGECELLS" })
    // exactly as long as its count: fine
    expect((await readOk(one([X.mergeCells([[0, 1, 0, 1]])]))).details[0]!.merges).toHaveLength(1)
    const records = Array.from({ length: 10 * BATCH_RECORDS }, (_, k) => X.blank(k % 60_000, 0, 0))
    const wb = await Effect.runPromise(Xls.read(X.xls(one(records))))
    const run = await countYields(() => Effect.runPromise(Workbook.sheetDetails(wb, 0)))
    expect(run.yields).toBeGreaterThanOrEqual(8)
  })

  it("a FORMULA shorter than its fixed 14 bytes is MalformedFile; 14 bytes read", async () => {
    const bytes = [...X.u16(0), ...X.u16(0), ...X.u16(0), ...X.f64(5)]
    expect(await errorOf(one([X.record(0x0006, bytes.slice(0, 13))]))).toMatchObject({ _tag: "MalformedFile", reason: "truncated record 0x6" })
    expect(await cellsOf([X.record(0x0006, bytes)])).toEqual([[0, 0, "Number", 5]])
  })

  it("records of a nested substream (BOF ... EOF) are skipped, the sheet's own after it count", async () => {
    const { details, sheets } = await readOk(one([
      X.colInfo(0, 0, true),
      X.bof(0x0020),
      X.colInfo(3, 3, true),
      X.mergeCells([[9, 9, 0, 1]]),
      X.dimensions(0, 9, 0, 9),
      X.number(0, 0, 0, 99),
      X.eof(),
      X.colInfo(5, 5, true),
      X.number(0, 1, 0, 1)
    ]))
    expect(details[0]).toEqual({ hiddenColumns: [{ first: 0, last: 0 }, { first: 5, last: 5 }], merges: [] })
    expect(values(sheets[0]!)).toEqual([[0, false, [[1, "Number", 1]]]])
  })

  it("a BIFF2-4 BOF record inside a sheet is an ordinary record: it opens no substream", async () => {
    // BOF record numbers of BIFF2 (0x0009), BIFF3 (0x0209) and BIFF4 (0x0409): only BIFF5/8's 0x0809 opens one
    for (const type of [0x0009, 0x0209, 0x0409]) {
      const old = X.record(type, [...X.u16(0), ...X.u16(0x0020)])
      const { details, sheets } = await readOk(one([X.number(0, 0, 0, 1), old, X.colInfo(5, 5, true), X.number(1, 0, 0, 2), X.eof(), X.number(2, 0, 0, 3)]))
      // the EOF after it is the sheet's own: nothing after it is read
      expect(details[0], type.toString(16)).toEqual({ hiddenColumns: [{ first: 5, last: 5 }], merges: [] })
      expect(values(sheets[0]!), type.toString(16)).toEqual([[0, false, [[0, "Number", 1]]], [1, false, [[0, "Number", 2]]]])
    }
  })

  it("an EOF inside a doubly nested substream closes one level only", async () => {
    // BOF BOF EOF (records of the outer nested substream) EOF: those records are skipped too
    const { details, sheets } = await readOk(one([
      X.number(0, 0, 0, 1),
      X.bof(0x0020),
      X.bof(0x0020),
      X.eof(),
      X.colInfo(3, 3, true),
      X.mergeCells([[9, 9, 0, 1]]),
      X.number(5, 0, 0, 99),
      X.eof(),
      X.colInfo(5, 5, true),
      X.number(1, 0, 0, 2)
    ]))
    expect(details[0]).toEqual({ hiddenColumns: [{ first: 5, last: 5 }], merges: [] })
    expect(values(sheets[0]!)).toEqual([[0, false, [[0, "Number", 1]]], [1, false, [[0, "Number", 2]]]])
    // the record-header scan nests the same way: row 1 still streams before a broken tail is read
    const nestedTwice = [X.bof(0x0020), X.bof(0x0020), X.eof(), X.number(5, 0, 0, 99), X.eof()]
    const filler = Array.from({ length: 2 * BATCH_RECORDS }, (_, k) => X.number(2 + k, 0, 0, k))
    const stream = X.workbookStream(one([X.number(0, 0, 0, 1), ...nestedTwice, X.number(1, 0, 0, 2), ...filler]))
    const wb = await Effect.runPromise(Xls.read(writeCfb([{ name: "Workbook", data: stream.subarray(0, stream.length - 10) }]).bytes))
    const first = await Effect.runPromise(Stream.runCollect(Stream.take(Workbook.rows(wb, 0), 2)))
    expect(Array.from(first).map((r) => r.index)).toEqual([0, 1])
  })

  it("a sheet offset that isn't at a BOF fails rows and sheetDetails", async () => {
    const stream = X.workbookStream(one([X.number(0, 0, 0, 1)]))
    const view = new DataView(stream.buffer)
    const sheetAt = stream.indexOf(0x85) // the BOUNDSHEET record's type
    const off = stream.slice()
    new DataView(off.buffer).setUint32(sheetAt + 4, view.getUint32(sheetAt + 4, true) + 4, true)
    const wb = await Effect.runPromise(Xls.read(writeCfb([{ name: "Workbook", data: off }]).bytes))
    const reason = "the sheet doesn't start with a BOF"
    expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, 0)))))).toMatchObject({ _tag: "MalformedFile", reason })
    expect(failureOf(await Effect.runPromise(Effect.exit(Workbook.sheetDetails(wb, 0))))).toMatchObject({ _tag: "MalformedFile", reason })
  })
})

/** A substream for `layout`: its records, as given (no EOF added). */
interface Sub {
  readonly name: string
  /** BOUNDSHEET dt: 0 worksheet, 2 chart, 6 VB module. */
  readonly dt?: number
  readonly records: ReadonlyArray<Uint8Array>
  /** The BOUNDSHEET offset instead of the substream's own (globals end and substream offsets given). */
  readonly at?: (globalsEnd: number, offsets: ReadonlyArray<number>) => number
}

const concat = (parts: ReadonlyArray<Uint8Array>) => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  // progress: one part per iteration
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

/**
 * A BIFF8 workbook stream: globals with one BOUNDSHEET per sub (in `listed`
 * order, default as given), then the substreams in the order given.
 */
const layout = (subs: ReadonlyArray<Sub>, listed: ReadonlyArray<number> = subs.map((_, k) => k)): Uint8Array => {
  const boundsheet = (s: Sub, offset: number) => X.record(0x0085, [...X.u32(offset), 0, s.dt ?? 0, ...X.shortString(s.name)])
  const bodies = subs.map((s) => concat(s.records))
  const globalsEnd = concat([X.bof(0x0005), ...listed.map((k) => boundsheet(subs[k]!, 0)), X.eof()]).length
  const offsets: Array<number> = []
  let at = globalsEnd
  // progress: one substream per iteration
  for (const body of bodies) {
    offsets.push(at)
    at += body.length
  }
  const listing = listed.map((k) => boundsheet(subs[k]!, subs[k]!.at?.(globalsEnd, offsets) ?? offsets[k]!))
  return concat([X.bof(0x0005), ...listing, X.eof(), ...bodies])
}
const fromStream = (stream: Uint8Array) => writeCfb([{ name: "Workbook", data: stream }]).bytes
const rowsOf = (wb: Workbook.Workbook, sheet: number, meter: Meter = noMeter) =>
  Effect.runPromise(Effect.exit(metered(meter, Stream.runCollect(Workbook.rows(wb, sheet)))))
const detailsOf = (wb: Workbook.Workbook, sheet: number, meter: Meter = noMeter) =>
  Effect.runPromise(Effect.exit(metered(meter, Workbook.sheetDetails(wb, sheet))))
const sheet = (name: string, records: ReadonlyArray<Uint8Array>, eof = true): Sub => ({ name, records: [X.bof(0x0010), ...records, ...(eof ? [X.eof()] : [])] })
/** Many records that are no cells (a filler record type the reader ignores). */
const filler = (n: number) => Array.from({ length: n }, () => X.record(0x0225, X.u16(0)))

describe("xls sheet: each sheet reads its own substream once", () => {
  it("two sheets at the same offset, or one inside the globals (offset 0 too), are MalformedFile", async () => {
    const same = layout([sheet("A", [X.number(0, 0, 0, 1)]), { ...sheet("B", []), at: (_, offsets) => offsets[0]! }])
    const shared = failureOf(await Effect.runPromise(Effect.exit(Xls.read(fromStream(same)))))
    expect(shared).toMatchObject({ _tag: "MalformedFile" })
    expect((shared as { reason: string }).reason).toMatch(/^sheets "A" and "B" both start at \d+$/)
    for (const at of [0, 4, -1]) {
      const inside = layout([{ ...sheet("A", []), at: (globalsEnd) => (at === -1 ? globalsEnd - 4 : at) }])
      const failure = failureOf(await Effect.runPromise(Effect.exit(Xls.read(fromStream(inside)))))
      expect(failure, String(at)).toMatchObject({ _tag: "MalformedFile" })
      expect((failure as { reason: string }).reason, String(at)).toContain("inside the workbook globals")
    }
  })

  it("BOUNDSHEETs out of offset order still read; each sheet ends where the next substream starts", async () => {
    // listed B, A while A's substream comes first; A has no EOF
    const stream = layout([sheet("A", [X.number(0, 0, 0, 1)], false), sheet("B", [X.number(0, 0, 0, 2)])], [1, 0])
    const wb = await Effect.runPromise(Xls.read(fromStream(stream)))
    expect(wb.sheets.map((s) => s.name)).toEqual(["B", "A"])
    const b = await rowsOf(wb, 0)
    expect(Exit.isSuccess(b) && Array.from(b.value).flatMap((r) => r.cells.map(Cell.value))).toEqual([2])
    expect(failureOf(await rowsOf(wb, 1))).toMatchObject({ _tag: "MalformedFile", reason: "the sheet has no EOF" })
    expect(failureOf(await detailsOf(wb, 1))).toMatchObject({ _tag: "MalformedFile", reason: "the sheet has no EOF" })
  })

  it("a sheet without EOF stops at the next sheet, a nested BOF without EOF at the sheet's end: not reading on", async () => {
    const big = filler(20_000)
    const cases: Array<[string, Sub]> = [
      ["no EOF", sheet("A", [X.number(0, 0, 0, 1)], false)],
      ["nested BOF without EOF", sheet("A", [X.number(0, 0, 0, 1), X.bof(0x0020), X.number(0, 1, 0, 1)])]
    ]
    for (const [label, a] of cases) {
      const wb = await Effect.runPromise(Xls.read(fromStream(layout([a, sheet("B", big)]))))
      const steps = counter("steps")
      expect(failureOf(await rowsOf(wb, 0, steps)), label).toMatchObject({ _tag: "MalformedFile", reason: "the sheet has no EOF" })
      expect(failureOf(await detailsOf(wb, 0, steps)), label).toMatchObject({ _tag: "MalformedFile", reason: "the sheet has no EOF" })
      // a few steps per record of A, none for B's 20 000
      expect(steps.total(), label).toBeLessThan(100)
    }
    // the last sheet ends at the stream's end
    const wb = await Effect.runPromise(Xls.read(fromStream(layout([sheet("A", [X.number(0, 0, 0, 1), X.bof(0x0020)])]))))
    expect(failureOf(await rowsOf(wb, 0))).toMatchObject({ reason: "the sheet has no EOF" })
  })

  it("a record running past its substream's end is MalformedFile", async () => {
    // A's last record claims 30 bytes: past A's end, into B
    const a = sheet("A", [X.number(0, 0, 0, 1), new Uint8Array([...X.u16(0x0225), ...X.u16(30), 0, 0])], false)
    const wb = await Effect.runPromise(Xls.read(fromStream(layout([a, sheet("B", filler(10))]))))
    expect(failureOf(await rowsOf(wb, 0))).toMatchObject({ _tag: "MalformedFile", reason: "record 0x225 runs past the end of its substream" })
    expect(failureOf(await detailsOf(wb, 0))).toMatchObject({ _tag: "MalformedFile", reason: "record 0x225 runs past the end of its substream" })
  })

  it("VB module offsets: a stale one (0) is ignored, one at a BOF past the globals ends the sheet before it", async () => {
    const module = (at: (globalsEnd: number, offsets: ReadonlyArray<number>) => number): Sub => ({ name: "Module1", dt: 6, records: [X.bof(0x0006), ...filler(5000), X.eof()], at })
    // listed: sheet A (no EOF), the module (its substream right after A), sheet B
    const stream = layout([sheet("A", [X.number(0, 0, 0, 1)], false), module((_, offsets) => offsets[1]!), sheet("B", [X.number(0, 0, 0, 2)])])
    const wb = await Effect.runPromise(Xls.read(fromStream(stream)))
    expect(wb.sheets.map((s) => s.name)).toEqual(["A", "B"])
    const steps = counter("steps")
    expect(failureOf(await rowsOf(wb, 0, steps))).toMatchObject({ reason: "the sheet has no EOF" })
    expect(steps.total()).toBeLessThan(100)
    // a module at offset 0 or inside the globals is no error and no substream start
    const stale = layout([sheet("A", [X.number(0, 0, 0, 1)]), { name: "Module1", dt: 6, records: [], at: () => 0 }, sheet("B", [X.number(0, 0, 0, 2)])])
    const ok = await Effect.runPromise(Xls.read(fromStream(stale)))
    const rows = await rowsOf(ok, 1)
    expect(Exit.isSuccess(rows) && Array.from(rows.value).flatMap((r) => r.cells.map(Cell.value))).toEqual([2])
  })

  it("many sheets cost at most the stream: total steps over all sheets' rows and details grow with its records", async () => {
    // 512 sheets without EOF of 20 records each: each used to walk on to the stream's end
    const subs = Array.from({ length: 512 }, (_, k) => sheet(`S${k}`, filler(20), false))
    const stream = layout(subs)
    const wb = await Effect.runPromise(Xls.read(fromStream(stream)))
    const records = 512 * 21
    const steps = counter("steps")
    // progress: one sheet per iteration
    for (let k = 0; k < 512; k++) {
      expect(failureOf(await rowsOf(wb, k, steps))).toMatchObject({ reason: "the sheet has no EOF" })
      expect(failureOf(await detailsOf(wb, k, steps))).toMatchObject({ reason: "the sheet has no EOF" })
    }
    // rows: the index pass (no row is complete), details: one pass (each step at most once per record)
    expect(steps.total()).toBeLessThanOrEqual(3 * records)
  })

  it("the rows reader's header pass yields in bounded batches before the first row", async () => {
    const wb = await Effect.runPromise(Xls.read(X.xls(one([X.number(0, 0, 0, 1), ...filler(10 * BATCH_RECORDS)]))))
    const run = await countYields(() => Effect.runPromise(Stream.runCollect(Stream.take(Workbook.rows(wb, 0), 1))))
    expect(run.yields).toBeGreaterThanOrEqual(9)
  })
})

describe("xls sheet: the rows reader's phases", () => {
  it("a substream that doesn't open fails every run of the rows, the same way", async () => {
    const stream = layout([{ ...sheet("A", [X.number(0, 0, 0, 1)]), at: (_, offsets) => offsets[0]! + 4 }])
    const wb = await Effect.runPromise(Xls.read(fromStream(stream)))
    const rows = Workbook.rows(wb, 0)
    for (let run = 0; run < 3; run++) {
      expect(failureOf(await Effect.runPromise(Effect.exit(Stream.runCollect(rows))))).toMatchObject({
        _tag: "MalformedFile",
        reason: "the sheet doesn't start with a BOF"
      })
    }
  })

  it("the header scan runs over several pulls, and the first row comes only once it has read the whole sheet", async () => {
    const wb = await Effect.runPromise(Xls.read(X.xls(one([X.number(0, 0, 0, 1), ...filler(3 * BATCH_RECORDS)]))))
    const steps = counter("steps")
    const run = await countYields(() => Effect.runPromise(metered(steps, Stream.runCollect(Stream.take(Workbook.rows(wb, 0), 1)))))
    expect(Array.from(run.value).map((r) => r.index)).toEqual([0])
    expect(run.yields).toBeGreaterThanOrEqual(3)
    expect(steps.total()).toBeGreaterThan(3 * BATCH_RECORDS)
  })

  it("rows left at the EOF beyond a batch's budget come out over several pulls, then the rows end", async () => {
    // rows 1 on wait for row 0, whose only cell comes last: a formula
    // waiting for a STRING that never comes. At the EOF row 0 gets "", and
    // all rows are due at once, more than one batch's budget (each row
    // costs its cells, being out of column order, plus one)
    const count = 12_000
    const records = [
      ...Array.from({ length: count }, (_, k) => [X.number(k + 1, 5, 0, k), X.mulBlank(k + 1, 0, [0, 0, 0, 0])]).flat(),
      X.formula(0, 0, 0, { kind: 0 })
    ]
    const { sheets } = await readOk(one(records))
    const rows = sheets[0]!
    expect(rows.length).toBe(count + 1)
    expect(values(rows.slice(0, 2))).toEqual([
      [0, false, [[0, "Text", ""]]],
      [1, false, [[0, "Empty", undefined], [1, "Empty", undefined], [2, "Empty", undefined], [3, "Empty", undefined], [5, "Number", 0]]]
    ])
    expect(rows.map((r) => r.index)).toEqual(Array.from({ length: count + 1 }, (_, r) => r))
  })

  it("no record is read after the sheet's EOF", async () => {
    // a record after the EOF that fails if read: a NUMBER too short for its cell
    const stream = layout([sheet("A", [X.number(0, 0, 0, 1), X.eof(), X.record(0x0203, X.u16(0))], false)])
    const wb = await Effect.runPromise(Xls.read(fromStream(stream)))
    const rows = await rowsOf(wb, 0)
    expect(Exit.isSuccess(rows) && Array.from(rows.value).map((r) => r.index)).toEqual([0])
  })

  it("the index pass counts one step per record it reads, the sheet's EOF included; the rows one per record, cell and row", async () => {
    // an empty sheet: the index pass reads its EOF; with an embedded
    // substream (BOF, EOF) before it, 3 records; a row of one NUMBER: the
    // index pass reads 2 records, the row's record is read again, its cell
    // and the row
    const cases: ReadonlyArray<[string, ReadonlyArray<Uint8Array>, number]> = [
      ["an empty sheet", [], 1],
      ["an embedded substream only", [X.bof(0x0020), X.eof()], 3],
      ["a row of one cell", [X.number(0, 0, 0, 1)], 5]
    ]
    for (const [label, records, expected] of cases) {
      const wb = await Effect.runPromise(Xls.read(X.xls(one(records))))
      const steps = counter("steps")
      const rows = await rowsOf(wb, 0, steps)
      expect(Exit.isSuccess(rows) && Array.from(rows.value).length, label).toBe(label === "a row of one cell" ? 1 : 0)
      expect(steps.total(), label).toBe(expected)
    }
  })
})

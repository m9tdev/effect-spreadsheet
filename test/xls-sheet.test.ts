/**
 * XLS worksheet substreams (src/internal/xls/sheet.ts): formula string
 * results across SHRFMLA/ARRAY/TABLE and CONTINUE, the BIFF4 FORMULA record
 * number, rows whose cells come in several places, non-finite numbers and
 * error codes, and sheetDetails edge cases. Synthetic records only.
 */
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { BATCH_RECORDS } from "../src/internal/xls/sheet.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
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
  })

  it("a continued STRING over maxStringLength, by its count or by its bytes, is LimitExceeded", async () => {
    const long = [X.record(STRING, [...X.u16(20), 0, ...chars("0123456789")]), X.record(CONTINUE, [0, ...chars("0123456789")])]
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), ...long]), { maxStringLength: 10 })).toMatchObject({ limit: "maxStringLength", max: 10 })
    expect(await cellsOf([X.formula(0, 0, 0, { kind: 0 }), ...long])).toEqual([[0, 0, "Text", "01234567890123456789"]])
    // a short count but CONTINUEs far over any string of maxStringLength characters
    const filler = Array.from({ length: 40 }, () => X.record(CONTINUE, new Array(8000).fill(0)))
    expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), X.record(STRING, [...X.u16(1), 0, 0x41]), ...filler]), { maxStringLength: 10 }))
      .toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
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

  it("a duplicate column in a row keeps the last cell", async () => {
    expect(await cellsOf([X.number(0, 3, 0, 1), X.number(0, 1, 0, 2), X.number(0, 3, 0, 3)])).toEqual([
      [0, 1, "Number", 2],
      [0, 3, "Number", 3]
    ])
    expect(await cellsOf([X.number(0, 0, 0, 1), X.number(0, 0, 0, 2)])).toEqual([[0, 0, "Number", 2]])
  })

  it("rows in order still stream: the first one comes out before the rest of the sheet is read", async () => {
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

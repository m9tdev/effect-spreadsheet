/**
 * XLS worksheet substreams (src/internal/xls/sheet.ts): formula string
 * results across SHRFMLA/ARRAY/TABLE and CONTINUE, the BIFF4 FORMULA record
 * number, rows whose cells come in several places, non-finite numbers and
 * error codes, and sheetDetails edge cases. Synthetic records only.
 */
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { BATCH_RECORDS, caps, Rows, type SheetContext, work } from "../src/internal/xls/sheet.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { writeCfb } from "./helpers/cfb.ts"
import { failureOf } from "./helpers/run.ts"
import { countYields } from "./helpers/yields.ts"
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

  it("a row whose last cell is a formula with a string result is emitted only after its STRING", async () => {
    const { sheets } = await readOk(one([X.number(0, 0, 0, 1), X.formula(0, 1, 0, { kind: 0 }), X.string("s"), X.number(1, 0, 0, 2)]))
    expect(values(sheets[0]!)).toEqual([
      [0, false, [[0, "Number", 1], [1, "Text", "s"]]],
      [1, false, [[0, "Number", 2]]]
    ])
  })

  it("rows held back by a row split far apart are capped: past the cap the sheet is MalformedFile", async () => {
    // a BLANK in row 0, rows 1-20 of 100 cells each, then row 0's second BLANK: rows 1-20 wait for row 0
    const records = [X.blank(0, 0, 0), ...Array.from({ length: 20 }, (_, r) => X.mulrk(r + 1, 0, Array.from({ length: 100 }, (_, c) => [0, X.rkInt(c)] as const))), X.blank(0, 1, 0)]
    const saved = caps.heldCells
    try {
      caps.heldCells = 1000
      const error = await errorOf(one(records))
      expect(error).toMatchObject({ _tag: "MalformedFile", reason: "rows split too far apart: over 1000 cells held for rows not complete" })
      // at 1 000 cells (row 0's first BLANK and rows 1-10), not before: the record that passes it
      expect((error as { offset: number }).offset).toBeGreaterThan(0)
      // all 2 002 cells are held when the last BLANK comes: a cap of 2 002 reads, 2 001 doesn't
      caps.heldCells = 2001
      expect(await errorOf(one(records))).toMatchObject({ reason: "rows split too far apart: over 2001 cells held for rows not complete" })
      caps.heldCells = 2002
      const { sheets } = await readOk(one(records))
      expect(sheets[0]!.map((r) => [r.index, r.cells.length])).toEqual([[0, 2], ...Array.from({ length: 20 }, (_, r) => [r + 1, 100])])
    } finally {
      caps.heldCells = saved
    }
  })

  it("held cells are exactly the cells of rows not emitted, through compaction and emitting", async () => {
    const limits = { ...SpreadsheetLimits.defaults, maxColumns: 8 }
    const rows = new Rows({ limits } as unknown as SheetContext)
    rows.start(new Map([[0, 10_000], [1, 20_000]]))
    const open = rows as unknown as { open: Map<number, { cells: Array<unknown> }> }
    const heldNow = () => [...open.open.values()].reduce((n, r) => n + r.cells.length, 0)
    // progress: k up to 100
    for (let k = 0; k < 100; k++) {
      expect(rows.cell(k % 2, Cell.number(7 - (k % 8), k))).toBeUndefined()
      expect(rows.held).toBe(heldNow())
    }
    expect(rows.complete(10_000)).toBeUndefined() // row 0 out
    expect(rows.held).toBe(heldNow())
    expect(rows.held).toBeGreaterThan(0)
    expect(rows.complete(20_000)).toBeUndefined()
    expect(rows.held).toBe(0)
  })

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
  it("complete(Infinity) with a string still pending stops once every row is out", () => {
    const rows = new Rows({ limits: SpreadsheetLimits.defaults } as unknown as SheetContext)
    rows.start(new Map([[0, 5]]))
    expect(rows.cell(0, Cell.number(0, 1))).toBeUndefined()
    expect(rows.complete(5)).toBeUndefined()
    // a FORMULA on a row the scan didn't place still waits for its STRING
    rows.pendingString = { row: 1, column: 0, meta: undefined }
    // the loop's bound is order.length: count the reads of `order`, so a
    // loop that only stops by luck fails here instead of hanging the worker
    const inner = (rows as unknown as { order: Uint32Array }).order
    let reads = 0
    ;(rows as unknown as { order: Uint32Array }).order = new Proxy(inner, {
      get: (target, key) => {
        if (++reads > 1000) throw new Error("complete() doesn't stop")
        const value = Reflect.get(target, key, target)
        return typeof value === "function" ? value.bind(target) : value
      }
    })
    expect(rows.complete(Infinity)).toBeUndefined()
    expect(rows.take().map((r) => r.index)).toEqual([0])
  })

  it("ROW-only rows count towards a batch: complete() and end() stop at the budget and go on in the next batch", () => {
    const rows = new Rows({ limits: SpreadsheetLimits.defaults } as unknown as SheetContext)
    // rows 0-999 declared by ROW records only, row 1000 with a cell (its last at offset 5), row 2000 declared after it
    rows.start(new Map([[1000, 5]]))
    // progress: r up to 1000
    for (let r = 0; r < 1000; r++) rows.row(r, r % 3 === 0)
    rows.row(2000, false)
    expect(rows.cell(1000, Cell.number(0, 1))).toBeUndefined()
    const out: Array<number> = []
    const drain = () => {
      // progress: each take() takes at least one row
      while (rows.queued()) out.push(...rows.take().map((r) => r.index))
    }
    // progress: each batch emits rows (budget 100 > 0), up to 1 002 batches
    for (let batch = 0; rows.behind(5); batch++) {
      if (batch > 1002) throw new Error("complete() doesn't go on")
      rows.budget = rows.effort + 100
      expect(rows.complete(5)).toBeUndefined()
      const before = out.length
      drain()
      // a budget of 100: at most 100 rows, one more for the row it ends with
      expect(out.length - before).toBeLessThanOrEqual(101)
    }
    expect(out).toEqual(Array.from({ length: 1001 }, (_, r) => r))
    // after the EOF: rows the scan didn't place and the declared rest, through end()
    expect(rows.cell(1500, Cell.number(0, 2))).toBeUndefined()
    // progress: r up to 1000
    for (let r = 0; r < 1000; r++) rows.row(3000 + r, false)
    let ended: unknown = false
    // progress: each batch emits rows (budget 100 > 0), up to 1 002 batches
    for (let batch = 0; ended !== true; batch++) {
      if (batch > 1002) throw new Error("end() doesn't go on")
      rows.budget = rows.effort + 100
      ended = rows.end()
      expect(typeof ended).toBe("boolean")
      const before = out.length
      drain()
      expect(out.length - before).toBeLessThanOrEqual(101)
    }
    expect(out.slice(1001)).toEqual([1500, 2000, ...Array.from({ length: 1000 }, (_, r) => 3000 + r)])
  })

  it("a repeated ROW record isn't queued again: a flood of one row costs one step to emit", () => {
    const rows = new Rows({ limits: SpreadsheetLimits.defaults } as unknown as SheetContext)
    rows.start(new Map())
    // progress: k up to 10 000
    for (let k = 0; k < 10_000; k++) rows.row(k % 2, k === 7)
    const before = work.steps
    expect(rows.end()).toBe(true)
    expect(work.steps - before).toBeLessThanOrEqual(2)
    expect(rows.take().map((r) => [r.index, r.hidden])).toEqual([[0, false], [1, true]])
  })
})

describe("xls sheet: repeated columns", () => {
  it("a row's repeated columns don't pile up: past 2 x maxColumns cells it is compacted, the last cell of a column kept", async () => {
    const limits = { ...SpreadsheetLimits.defaults, maxColumns: 8 }
    const rows = new Rows({ limits } as unknown as SheetContext)
    rows.start(new Map([[0, Infinity]]))
    let most = 0
    // 1 000 passes over columns 7..0 (unsorted): 8 000 cells for 8 columns
    for (let k = 0; k < 8000; k++) {
      expect(rows.cell(0, Cell.number(7 - (k % 8), k))).toBeUndefined()
      most = Math.max(most, (rows as unknown as { open: Map<number, { cells: Array<unknown> }> }).open.get(0)!.cells.length)
    }
    expect(most).toBeLessThanOrEqual(2 * 8 + 1)
    expect(rows.end()).toBe(true)
    const [row] = rows.take()
    expect(row.cells.map((c) => [c.column, Cell.value(c)])).toEqual(Array.from({ length: 8 }, (_, c) => [c, 8000 - 8 + 7 - c]))
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
const rowsOf = (wb: Workbook.Workbook, sheet: number) => Effect.runPromise(Effect.exit(Stream.runCollect(Workbook.rows(wb, sheet))))
const detailsOf = (wb: Workbook.Workbook, sheet: number) => Effect.runPromise(Effect.exit(Workbook.sheetDetails(wb, sheet)))
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
      const before = work.steps
      expect(failureOf(await rowsOf(wb, 0)), label).toMatchObject({ _tag: "MalformedFile", reason: "the sheet has no EOF" })
      expect(failureOf(await detailsOf(wb, 0)), label).toMatchObject({ _tag: "MalformedFile", reason: "the sheet has no EOF" })
      // a few steps per record of A, none for B's 20 000
      expect(work.steps - before, label).toBeLessThan(100)
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
    const before = work.steps
    expect(failureOf(await rowsOf(wb, 0))).toMatchObject({ reason: "the sheet has no EOF" })
    expect(work.steps - before).toBeLessThan(100)
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
    const before = work.steps
    // progress: one sheet per iteration
    for (let k = 0; k < 512; k++) {
      expect(failureOf(await rowsOf(wb, k))).toMatchObject({ reason: "the sheet has no EOF" })
      expect(failureOf(await detailsOf(wb, k))).toMatchObject({ reason: "the sheet has no EOF" })
    }
    // rows: the header pass and the main pass, details: one pass (each step at most once per record)
    expect(work.steps - before).toBeLessThanOrEqual(3 * records)
  })

  it("the rows reader's header pass yields in bounded batches before the first row", async () => {
    const wb = await Effect.runPromise(Xls.read(X.xls(one([X.number(0, 0, 0, 1), ...filler(10 * BATCH_RECORDS)]))))
    const run = await countYields(() => Effect.runPromise(Stream.runCollect(Stream.take(Workbook.rows(wb, 0), 1))))
    expect(run.yields).toBeGreaterThanOrEqual(9)
  })
})

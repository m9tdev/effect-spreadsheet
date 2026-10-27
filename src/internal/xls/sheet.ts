/**
 * A worksheet substream of a BIFF8/BIFF5 workbook stream -> rows, streamed.
 *
 * The substream starts at its BOUNDSHEET offset with a worksheet BOF and
 * runs to its EOF. Records are read in batches (at most `BATCH_RECORDS`
 * records or about `BATCH_CELLS` cells per pull, with a yield before each),
 * and rows come out in chunks of at most `BATCH_RECORDS` rows or about
 * `BATCH_CELLS` cells, so the work between event-loop turns is bounded.
 *
 * - Rows: ROW records give hidden rows (fDyZero) and rows without cells;
 *   cells mostly come row by row after their block of ROW records, but
 *   some writers put a row's cells in several places. A first pass over
 *   the record headers (`RowScan`) finds where each row's last cell is; a
 *   row is emitted, merged, once that record is handled, in row order
 *   (rows before it wait, holding only cells already counted against
 *   `maxCellsPerSheet`). ROW-only rows are emitted in order between them.
 *   Hidden rows are emitted with `hidden: true`, never skipped.
 * - Cells: LABELSST, LABEL, RSTRING (text), NUMBER, RK, MULRK (numbers),
 *   BOOLERR (booleans and errors), BLANK, MULBLANK (Empty), FORMULA with
 *   its cached result (a following STRING record for text, after any
 *   SHRFMLA/ARRAY/TABLE, maybe continued by CONTINUE records). NaN and
 *   infinities are `#NUM!` errors; unknown error codes are `#N/A`. Numbers whose
 *   XF has a date format become Dates, as in XLSX. Formula text isn't
 *   decompiled (no `formula`).
 * - Nothing is sized by DIMENSIONS; cost follows the records present.
 *
 * @internal
 */
import { Cause, Effect, Stream } from "effect"
import type * as Arr from "effect/Array"
import type * as Pull from "effect/Pull"
import * as Cell from "../../Cell.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { ColumnSpan, Range, Row, SheetDetails } from "../../Workbook.ts"
import { BOF_BIFF5_8, type BiffRecord, type ContinuedRecord, EOF, RecordReader, readBof } from "../biff/records.ts"
import { serialToDate } from "../xlsx/numFmt.ts"
import { readRichString, SegmentReader } from "../biff/segments.ts"
import { decodeRk, ERROR_CODES, Fields, readString } from "./fields.ts"
import type { Globals } from "./globals.ts"

/**
 * Steps of the rows reader's loops (records scanned and read, cells, rows,
 * queue and sort steps) and of sheetDetails' record walk, summed over every read in this process. Tests use
 * it to check that the work grows linearly with the sheet, whatever the
 * machine's load; it costs one increment per step.
 */
export const work = { steps: 0 }

/** Records handled per pull: bounded work between event-loop turns. */
export const BATCH_RECORDS = 4096
/**
 * Cells per pull, read or emitted: a batch ends once it passes this (one
 * MULBLANK holds up to 4 109 cells), and so does a chunk of rows (a row
 * holds at most `maxColumns` cells, see `compact`).
 */
export const BATCH_CELLS = 65_536

/**
 * Cells held in rows not emitted yet (rows waiting for a later cell of
 * theirs, or for an earlier row): at most `heldCells`, 16 batches. Rows
 * stream in row order, so a row whose cells come both early and late in
 * the sheet holds back every row after it; real files split rows across a
 * few rows only (the corpus holds at most a few thousand cells). Past it,
 * the sheet is MalformedFile: MS-XLS puts a row's cells in its row block.
 * Tests may lower it.
 */
export const caps = { heldCells: 16 * BATCH_CELLS }

const ROW = 0x0208
const LABELSST = 0x00FD
const LABEL = 0x0204
const RSTRING = 0x00D6
const NUMBER = 0x0203
const RK = 0x027E
const MULRK = 0x00BD
const BOOLERR = 0x0205
const BLANK = 0x0201
const MULBLANK = 0x00BE
const FORMULA = 0x0006
/** FORMULA with the BIFF3/BIFF4 record number, which some writers still use: same fields up to the result. */
const FORMULA_BIFF4 = 0x0406
const STRING = 0x0207
/** May sit between a FORMULA and its STRING (MS-XLS 2.1.7.20.5: FORMULA [SHRFMLA / ARRAY / TABLE] [STRING]). */
const SHRFMLA = 0x04BC
const ARRAY = 0x0221
const TABLE = 0x0236
const DIMENSIONS = 0x0200
const COLINFO = 0x007D
const MERGECELLS = 0x00E5

export interface SheetContext {
  readonly stream: Uint8Array
  readonly offset: number
  /** End of the substream: reading stops there (see `SheetEntry.end`). */
  readonly end: number
  readonly part: string
  readonly globals: Globals
  readonly limits: Limits
}

const malformed = (part: string, reason: string, offset?: number) =>
  new MalformedFile({ format: "xls", part, reason, ...(offset === undefined ? {} : { offset }) })

/**
 * The data of a STRING record and its CONTINUEs may take at most this many
 * bytes: the header (9), 2 bytes per character plus a 5-byte CONTINUE
 * header and option byte per character in the worst split, and 4 bytes
 * per rich-text run. More is a string over `maxStringLength`.
 */
const stringMaxBytes = (limits: Limits): number => 16 + 7 * limits.maxStringLength + 4 * 0xFFFF

/**
 * Unknown error codes (BOOLERR, FORMULA results) give `#N/A`, not a failure:
 * as the XLSX reader, which keeps whatever error text a cell has.
 */
const errorCode = (code: number): string => ERROR_CODES.get(code) ?? "#N/A"

/** Opens the substream: checks the offset and the worksheet BOF; the reader after it, or why not. */
const openSubstream = (context: SheetContext): RecordReader | SpreadsheetError => {
  if (context.offset < 0 || context.offset + 4 > context.stream.length) {
    return malformed(context.part, `sheet offset ${context.offset} is outside the workbook stream`)
  }
  const reader = new RecordReader(context.stream, context.part, context.offset, context.end)
  const bof = reader.next()
  if (bof === undefined) return malformed(context.part, "the sheet has no BOF", context.offset)
  if (bof instanceof MalformedFile) return bof
  const parsed = readBof(bof)
  if (parsed === undefined) return malformed(context.part, "the sheet doesn't start with a BOF", context.offset)
  return reader
}

/** Cell records: each starts with its row (u16). */
const CELL_RECORDS: ReadonlySet<number> = new Set([LABELSST, LABEL, RSTRING, NUMBER, RK, MULRK, BOOLERR, BLANK, MULBLANK, FORMULA, FORMULA_BIFF4])

/**
 * A first pass over the substream's records (headers and row numbers only):
 * the offset of each row's last cell record. Some writers emit a row's
 * cells in several places; knowing where each row ends lets the main pass
 * merge them and still emit every row as soon as it is complete. One entry
 * per row with cells (at most 65 536: rows are u16).
 */
class RowScan {
  readonly lastAt = new Map<number, number>()
  done = false
  private nested = 0
  private readonly reader: RecordReader

  constructor(context: SheetContext, offset: number) {
    this.reader = new RecordReader(context.stream, context.part, offset, context.end)
  }

  /** Reads up to `max` records; `done` at the sheet's EOF or where the records stop (the main pass reports why). */
  step(max: number): void {
    // progress: k up to max; every next() advances the reader
    for (let k = 0; k < max; k++) {
      work.steps++
      const record = this.reader.next()
      if (record === undefined || record instanceof MalformedFile) {
        this.done = true
        return
      }
      // nested substreams as the main pass sees them (see `handle`)
      if (record.type === BOF_BIFF5_8) this.nested++
      else if (record.type === EOF) {
        if (this.nested === 0) {
          this.done = true
          return
        }
        this.nested--
      } else if (this.nested === 0 && CELL_RECORDS.has(record.type) && record.data.length >= 2) {
        this.lastAt.set(record.data[0]! | (record.data[1]! << 8), record.offset)
      }
    }
  }
}

/**
 * Sorts a row's cells by column in place, keeping the last cell of each
 * column (`sort` is stable), and returns them.
 */
const compact = (cells: Array<Cell.Cell>): Array<Cell.Cell> => {
  cells.sort((a, b) => a.column - b.column)
  let kept = 0
  // progress: k up to cells.length
  for (let k = 0; k < cells.length; k++) {
    work.steps++
    const next = cells[k + 1]
    if (next === undefined || next.column !== cells[k]!.column) cells[kept++] = cells[k]!
  }
  cells.length = kept
  return cells
}

interface OpenRow {
  cells: Array<Cell.Cell>
  /** Columns came in increasing order so far. */
  sorted: boolean
  lastColumn: number
}

/**
 * Builds rows from cell records. Cells of a row are gathered until the
 * record holding its last cell (from the `RowScan`) is handled, then rows
 * are emitted in row order. Rows waiting for later cells hold only cells
 * already counted against `maxCellsPerSheet`.
 */
export class Rows {
  /** Rows ready to emit; `out[outAt]` is the next one (see `take`). */
  private out: Array<Row> = []
  private outAt = 0
  /** Rows with cells, not emitted yet. */
  private readonly open = new Map<number, OpenRow>()
  /** The row the last cell went to (most cells follow the one before). */
  private hotRow = -1
  private hotEntry: OpenRow | undefined
  /** Offset of each row's last cell record. */
  private lastAt: ReadonlyMap<number, number> = new Map()
  /** Rows with cells, ascending; `order[nextRow]` is the next one to emit. */
  private order: Uint32Array = new Uint32Array(0)
  private nextRow = 0
  /** Offset of `order[nextRow]`'s last cell record (Infinity when none is left). */
  private nextAt = Infinity
  /** Hidden flags from ROW records. */
  private readonly hidden = new Set<number>()
  /** Rows declared by ROW records, not emitted yet, in record order. */
  private declared: Array<number> = []
  private declaredAt = 0
  /** Rows emitted so far (a ROW-only row is emitted once). */
  private readonly emitted = new Set<number>()
  private rowCount = 0
  /** Cells read so far (repeats included): counted against maxCellsPerSheet. */
  cellCount = 0
  /** Cells in `open` rows: exactly those held, not emitted yet (see `caps`). */
  held = 0
  /**
   * Cells read plus cells sorted (compacting a row): the work that a batch
   * spends, up to `budget`. Emitting rows stops at the budget and goes on
   * in the next batch.
   */
  effort = 0
  budget = Infinity
  /** Depth of substreams nested in this one (embedded charts: BOF ... EOF). */
  nested = 0
  /** A FORMULA whose text result comes in the next STRING record. */
  pendingString: { column: number; row: number; meta: Cell.Meta | undefined } | undefined

  private readonly context: SheetContext

  constructor(context: SheetContext) {
    this.context = context
  }

  /** Where each row's cells end, from the `RowScan`. */
  start(lastAt: ReadonlyMap<number, number>): void {
    this.lastAt = lastAt
    this.order = Uint32Array.from(lastAt.keys()).sort()
    this.nextRow = 0
    this.nextAt = this.order.length > 0 ? lastAt.get(this.order[0]!)! : Infinity
  }

  private push(index: number, hidden: boolean, cells: Array<Cell.Cell>): SpreadsheetError | undefined {
    const limits = this.context.limits
    if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
    if (++this.rowCount > limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: this.rowCount })
    this.out.push({ index, hidden, cells })
    this.emitted.add(index)
    return undefined
  }

  /** Emits declared ROW-only rows up to (not including) `before`. */
  private flushDeclared(before: number): SpreadsheetError | undefined {
    // progress: declaredAt increases up to declared.length
    while (this.declaredAt < this.declared.length && this.declared[this.declaredAt]! < before) {
      work.steps++
      const index = this.declared[this.declaredAt++]!
      if (this.emitted.has(index)) continue
      const error = this.push(index, this.hidden.has(index), [])
      if (error !== undefined) return error
    }
    if (this.declaredAt === this.declared.length) {
      this.declared = []
      this.declaredAt = 0
    }
    return undefined
  }

  /** Emits row `index` with its cells (if it got any). */
  private finishRow(index: number): SpreadsheetError | undefined {
    const entry = this.open.get(index)
    if (entry === undefined) return undefined
    this.open.delete(index)
    this.held -= entry.cells.length
    if (this.hotRow === index) {
      this.hotRow = -1
      this.hotEntry = undefined
    }
    if (!entry.sorted) this.effort += entry.cells.length
    return this.push(index, this.hidden.has(index), entry.sorted ? entry.cells : compact(entry.cells))
  }

  /** Whether rows wait to be taken. */
  queued(): boolean {
    return this.outAt < this.out.length
  }

  /** The next chunk of rows: at least one, at most `BATCH_RECORDS` rows or until `BATCH_CELLS` cells are passed. */
  take(): Arr.NonEmptyReadonlyArray<Row> {
    const from = this.outAt
    let cells = 0
    // progress: outAt increases up to out.length
    do {
      work.steps++
      cells += this.out[this.outAt++]!.cells.length
    } while (this.outAt < this.out.length && this.outAt - from < BATCH_RECORDS && cells < BATCH_CELLS)
    const chunk = this.out.slice(from, this.outAt) as unknown as Arr.NonEmptyReadonlyArray<Row>
    if (this.outAt === this.out.length) {
      this.out = []
      this.outAt = 0
    }
    return chunk
  }

  /**
   * After the record at `offset`: emits, in row order, the rows whose last
   * cell record is at or before it (not one still waiting for its STRING),
   * until `effort` reaches `budget` (see `behind`).
   */
  complete(offset: number): SpreadsheetError | undefined {
    // progress: nextRow increases, bounded by order.length (offset may be Infinity)
    while (this.nextRow < this.order.length && this.nextAt <= offset && this.effort < this.budget) {
      work.steps++
      const index = this.order[this.nextRow]!
      if (this.pendingString?.row === index) return undefined
      this.nextRow++
      this.nextAt = this.nextRow < this.order.length ? this.lastAt.get(this.order[this.nextRow]!)! : Infinity
      const error = this.flushDeclared(index) ?? this.finishRow(index)
      if (error !== undefined) return error
    }
    return undefined
  }

  /** Whether rows complete at `offset` are left for `complete` (it stopped at the budget). */
  behind(offset: number): boolean {
    return this.nextRow < this.order.length && this.nextAt <= offset && this.pendingString?.row !== this.order[this.nextRow]
  }

  row(index: number, hidden: boolean): void {
    if (hidden) this.hidden.add(index)
    this.declared.push(index)
  }

  cell(row: number, cell: Cell.Cell): SpreadsheetError | undefined {
    work.steps++
    const limits = this.context.limits
    if (cell.column >= limits.maxColumns) {
      return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: cell.column + 1 })
    }
    if (++this.cellCount > limits.maxCellsPerSheet) {
      return new LimitExceeded({ limit: "maxCellsPerSheet", max: limits.maxCellsPerSheet, actual: this.cellCount })
    }
    let entry = this.hotRow === row ? this.hotEntry : this.open.get(row)
    if (entry === undefined) {
      entry = { cells: [], sorted: true, lastColumn: -1 }
      this.open.set(row, entry)
    }
    this.hotRow = row
    this.hotEntry = entry
    if (cell.column <= entry.lastColumn) entry.sorted = false
    entry.lastColumn = cell.column
    entry.cells.push(cell)
    this.held++
    this.effort++
    // repeated columns don't pile up: past 2 x maxColumns cells a row is
    // compacted (to at most maxColumns), so the work is amortized
    if (!entry.sorted && entry.cells.length > 2 * limits.maxColumns) {
      this.effort += entry.cells.length
      this.held -= entry.cells.length
      compact(entry.cells)
      this.held += entry.cells.length
      entry.sorted = true
      entry.lastColumn = entry.cells[entry.cells.length - 1]!.column
    }
    return undefined
  }

  /** After the EOF, once `complete(Infinity)` isn't `behind`: the rest. */
  end(): SpreadsheetError | undefined {
    // rows the scan didn't place (none when both passes agree), in row order
    const left = Uint32Array.from(this.open.keys()).sort()
    // progress: one row per iteration
    for (const index of left) {
      work.steps++
      const failed = this.flushDeclared(index) ?? this.finishRow(index)
      if (failed !== undefined) return failed
    }
    return this.flushDeclared(Infinity)
  }
}

/** A number cell, a Date when its XF has a date format. */
const numberCell = (globals: Globals, column: number, xf: number, value: number): Cell.Cell => {
  const format = globals.formats[xf]
  const meta: Cell.Meta | undefined = format?.code === undefined ? undefined : { numFmt: format.code }
  if (format?.date === true) {
    const date = serialToDate(value, globals.date1904)
    if (date !== undefined) return Cell.date(column, date, meta)
  }
  // NaN and infinities aren't Excel numbers: #NUM!, as for a formula result
  if (!Number.isFinite(value)) return Cell.error(column, "#NUM!", meta)
  return Cell.number(column, value === 0 ? 0 : value, meta)
}

/** Handles one record of the substream; `true` at its EOF. */
const handle = (record: BiffRecord | ContinuedRecord, rows: Rows, context: SheetContext): SpreadsheetError | boolean => {
  const { globals, part } = context
  // embedded substreams (chart objects) sit between BOF and EOF inside the sheet: skipped
  if (record.type === BOF_BIFF5_8) {
    rows.nested++
    return false
  }
  if (rows.nested > 0) {
    if (record.type === EOF) rows.nested--
    return false
  }
  const f = new Fields(record.data)
  const truncated = () => malformed(part, `truncated record 0x${record.type.toString(16)}`, record.offset)
  const tooLong = () => new LimitExceeded({ limit: "maxStringLength", max: context.limits.maxStringLength, actual: context.limits.maxStringLength + 1 })
  // a formula's string result comes right after it, or after its SHRFMLA, ARRAY or TABLE
  if (
    rows.pendingString !== undefined && record.type !== STRING && record.type !== SHRFMLA && record.type !== ARRAY &&
    record.type !== TABLE
  ) {
    const pending = rows.pendingString
    rows.pendingString = undefined
    const error = rows.cell(pending.row, Cell.text(pending.column, "", pending.meta))
    if (error !== undefined) return error
  }
  const rc = (): readonly [number, number, number] | undefined => {
    const row = f.u16(0)
    const col = f.u16(2)
    const xf = f.u16(4)
    return row === undefined || col === undefined || xf === undefined ? undefined : [row, col, xf]
  }
  switch (record.type) {
    case EOF:
      return true
    case ROW: {
      const row = f.u16(0)
      const grbit = f.u16(12)
      if (row === undefined || grbit === undefined) return truncated()
      rows.row(row, (grbit & 0x20) !== 0)
      return false
    }
    case LABELSST: {
      const at = rc()
      const isst = f.u32(6)
      if (at === undefined || isst === undefined) return truncated()
      const value = globals.strings[isst]
      if (value === undefined) return malformed(part, `shared string ${isst} doesn't exist`, record.offset)
      return rows.cell(at[0], Cell.text(at[1], value)) ?? false
    }
    case LABEL:
    case RSTRING: {
      const at = rc()
      if (at === undefined) return truncated()
      const text = readString(record.data, 6, globals.text, 2)
      if (text === undefined) return truncated()
      if (text === null) return tooLong()
      return rows.cell(at[0], Cell.text(at[1], text.value)) ?? false
    }
    case NUMBER: {
      const at = rc()
      const value = f.f64(6)
      if (at === undefined || value === undefined) return truncated()
      return rows.cell(at[0], numberCell(globals, at[1], at[2], value)) ?? false
    }
    case RK: {
      const at = rc()
      const rk = f.u32(6)
      if (at === undefined || rk === undefined) return truncated()
      return rows.cell(at[0], numberCell(globals, at[1], at[2], decodeRk(rk))) ?? false
    }
    case MULRK:
    case MULBLANK: {
      const row = f.u16(0)
      const first = f.u16(2)
      const last = f.u16(record.data.length - 2)
      const width = record.type === MULRK ? 6 : 2
      if (row === undefined || first === undefined || last === undefined || record.data.length < 6) return truncated()
      const count = (record.data.length - 6) / width
      if (!Number.isInteger(count) || last - first + 1 !== count) {
        return malformed(part, `record 0x${record.type.toString(16)} doesn't match its column range`, record.offset)
      }
      // progress: k up to count (bounded by the record length)
      for (let k = 0; k < count; k++) {
        const xf = f.u16(4 + width * k)!
        const cell = record.type === MULRK
          ? numberCell(globals, first + k, xf, decodeRk(f.u32(6 + width * k)!))
          : Cell.empty(first + k)
        const error = rows.cell(row, cell)
        if (error !== undefined) return error
      }
      return false
    }
    case BOOLERR: {
      const at = rc()
      const value = f.u8(6)
      const isError = f.u8(7)
      if (at === undefined || value === undefined || isError === undefined) return truncated()
      if (isError === 1) {
        return rows.cell(at[0], Cell.error(at[1], errorCode(value))) ?? false
      }
      return rows.cell(at[0], Cell.boolean(at[1], value !== 0)) ?? false
    }
    case BLANK: {
      const at = rc()
      if (at === undefined) return truncated()
      return rows.cell(at[0], Cell.empty(at[1])) ?? false
    }
    case FORMULA:
    case FORMULA_BIFF4: {
      const at = rc()
      if (at === undefined || record.data.length < 14) return truncated()
      if (f.u16(12) === 0xFFFF) {
        // a special result: string (in the next STRING record), boolean, error, empty string
        const kind = f.u8(6)!
        const value = f.u8(8)!
        if (kind === 0) {
          rows.pendingString = { row: at[0], column: at[1], meta: undefined }
          return false
        }
        if (kind === 1) return rows.cell(at[0], Cell.boolean(at[1], value !== 0)) ?? false
        if (kind === 2) return rows.cell(at[0], Cell.error(at[1], errorCode(value))) ?? false
        return rows.cell(at[0], Cell.text(at[1], "")) ?? false
      }
      const value = f.f64(6)!
      return rows.cell(at[0], numberCell(globals, at[1], at[2], value)) ?? false
    }
    case STRING: {
      const pending = rows.pendingString
      if (pending === undefined) return false // a STRING without a formula: ignored
      rows.pendingString = undefined
      // BIFF8: an XLUnicodeString, maybe continued (each CONTINUE restarts
      // with an option byte); BIFF5: a u16-count code-page string
      const read = "segments" in record
        ? readRichString(new SegmentReader(record.segments), context.limits.maxStringLength)
        : readString(record.data, 0, globals.text, 2)
      if (read === undefined) return truncated()
      if (read === null) return tooLong()
      const text = typeof read === "string" ? read : read.value
      return rows.cell(pending.row, Cell.text(pending.column, text, pending.meta)) ?? false
    }
    default:
      return false
  }
}

/** The rows of a worksheet substream. Re-runnable. */
export const sheetRows = (context: SheetContext): Stream.Stream<Row, SpreadsheetError> =>
  Stream.fromPull(Effect.sync(() => {
    const opened = openSubstream(context)
    const reader = opened instanceof RecordReader ? opened : undefined
    const scan = reader === undefined ? undefined : new RowScan(context, reader.offset)
    const rows = new Rows(context)
    const maxStringBytes = stringMaxBytes(context.limits)
    let done = false
    /** The EOF was handled: only rows are left to emit. */
    let ended = false
    /** Offset of the last record handled (rows complete up to it may wait for budget). */
    let handledAt = -1
    const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, SpreadsheetError> = Effect.gen(function*() {
      if (reader === undefined || scan === undefined) return yield* Effect.fail(opened as SpreadsheetError)
      // progress: each iteration takes queued rows, or reads up to
      // BATCH_RECORDS records (the scan's or the reader's offset grows with
      // each), or ends the stream
      while (true) {
        yield* Effect.yieldNow
        if (rows.queued()) return rows.take()
        if (done) return yield* Cause.done()
        if (!scan.done) {
          scan.step(BATCH_RECORDS)
          if (scan.done) rows.start(scan.lastAt)
          continue
        }
        rows.budget = rows.effort + BATCH_CELLS
        // rows a budget left complete but not emitted come first (after the
        // EOF, all the rest): each call emits at least one, or none is left
        const due = ended ? Infinity : handledAt
        const caught = rows.complete(due)
        if (caught !== undefined) return yield* Effect.fail(caught)
        if (ended) {
          if (!rows.behind(Infinity)) {
            done = true
            const error = rows.end()
            if (error !== undefined) return yield* Effect.fail(error)
          }
          if (rows.queued()) return rows.take()
          continue
        }
        // progress: k up to BATCH_RECORDS; every next() advances the reader
        for (let k = 0; k < BATCH_RECORDS && rows.effort < rows.budget && !rows.behind(handledAt); k++) {
          work.steps++
          const record = rows.nested === 0 && context.globals.biff8 && reader.peekType() === STRING
            ? reader.nextWithContinues(maxStringBytes, "maxStringLength")
            : reader.next()
          if (record === undefined) return yield* Effect.fail(malformed(context.part, "the sheet has no EOF"))
          if (record instanceof MalformedFile) return yield* Effect.fail(record)
          if (record instanceof LimitExceeded) {
            const max = context.limits.maxStringLength
            return yield* Effect.fail(new LimitExceeded({ limit: "maxStringLength", max, actual: max + 1 }))
          }
          const result = handle(record, rows, context)
          if (result === true) {
            ended = true
            break
          }
          if (result !== false) return yield* Effect.fail(result)
          if (rows.held > caps.heldCells) {
            return yield* Effect.fail(malformed(context.part, `rows split too far apart: over ${caps.heldCells} cells held for rows not complete`, record.offset))
          }
          handledAt = record.offset
          const error = rows.complete(record.offset)
          if (error !== undefined) return yield* Effect.fail(error)
        }
        if (rows.queued()) return rows.take()
      }
    })
    return pull
  }))

/** Merges sorted spans that overlap or touch. */
const mergeSpans = (spans: ReadonlyArray<ColumnSpan>): Array<ColumnSpan> => {
  const sorted = [...spans].sort((a, b) => a.first - b.first)
  const out: Array<{ first: number; last: number }> = []
  // progress: one span per iteration
  for (const span of sorted) {
    const last = out[out.length - 1]
    if (last !== undefined && span.first <= last.last + 1) last.last = Math.max(last.last, span.last)
    else out.push({ first: span.first, last: span.last })
  }
  return out
}

/**
 * Hidden columns (COLINFO), merges (MERGECELLS) and the declared dimension
 * (DIMENSIONS) of a worksheet substream, in one pass over its records
 * (yielding every `BATCH_RECORDS` records).
 */
export const sheetDetails = (context: SheetContext): Effect.Effect<SheetDetails, SpreadsheetError> =>
  Effect.gen(function*() {
    const opened = openSubstream(context)
    if (!(opened instanceof RecordReader)) return yield* Effect.fail(opened)
    const { part, limits, globals } = context
    const hidden: Array<ColumnSpan> = []
    const merges: Array<Range> = []
    let dimension: Range | undefined
    let colInfos = 0
    let nested = 0
    // progress: every next() advances the reader
    for (let k = 1;; k++) {
      work.steps++
      if (k % BATCH_RECORDS === 0) yield* Effect.yieldNow
      const record = opened.next()
      if (record === undefined) return yield* Effect.fail(malformed(part, "the sheet has no EOF"))
      if (record instanceof MalformedFile) return yield* Effect.fail(record)
      if (record.type === BOF_BIFF5_8) {
        nested++
        continue
      }
      if (record.type === EOF) {
        if (nested === 0) break
        nested--
        continue
      }
      if (nested > 0) continue
      const f = new Fields(record.data)
      if (record.type === DIMENSIONS) {
        const [r1, r2, c1, c2] = globals.biff8 ? [f.u32(0), f.u32(4), f.u16(8), f.u16(10)] : [f.u16(0), f.u16(2), f.u16(4), f.u16(6)]
        if (r1 !== undefined && r2 !== undefined && c1 !== undefined && c2 !== undefined && r2 > r1 && c2 > c1) {
          dimension = { firstRow: r1, firstColumn: c1, lastRow: r2 - 1, lastColumn: c2 - 1 }
        }
      } else if (record.type === COLINFO) {
        if (++colInfos > limits.maxColumns) {
          return yield* Effect.fail(new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: colInfos }))
        }
        const first = f.u16(0)
        const last = f.u16(2)
        const grbit = f.u16(8)
        if (first === undefined || last === undefined || grbit === undefined) {
          return yield* Effect.fail(malformed(part, "truncated COLINFO", record.offset))
        }
        // spans starting past the last column are ignored, as in XLSX
        if ((grbit & 0x01) !== 0 && last >= first && first < limits.maxColumns) hidden.push({ first, last: Math.min(last, limits.maxColumns - 1) })
      } else if (record.type === MERGECELLS) {
        const count = f.u16(0)
        if (count === undefined || 2 + 8 * count > record.data.length) {
          return yield* Effect.fail(malformed(part, "truncated MERGECELLS", record.offset))
        }
        // progress: j up to count (bounded by the record length)
        for (let j = 0; j < count; j++) {
          if (merges.length >= limits.maxCellsPerSheet) {
            return yield* Effect.fail(new LimitExceeded({ limit: "maxCellsPerSheet", max: limits.maxCellsPerSheet, actual: merges.length + 1 }))
          }
          const at = 2 + 8 * j
          const [r1, r2, c1, c2] = [f.u16(at)!, f.u16(at + 2)!, f.u16(at + 4)!, f.u16(at + 6)!]
          merges.push({ firstRow: Math.min(r1, r2), firstColumn: Math.min(c1, c2), lastRow: Math.max(r1, r2), lastColumn: Math.max(c1, c2) })
        }
      }
    }
    const details: SheetDetails = { hiddenColumns: mergeSpans(hidden), merges, ...(dimension === undefined ? {} : { dimension }) }
    return details
  })

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
 *   some writers put a row's cells in several places, or ROW records after
 *   the cells. A first pass over the record headers (`RowScan`) finds where
 *   each row's last cell is, and the rows ROW records declare or hide; a
 *   row is emitted, merged, once its last cell record is handled, in row
 *   order (rows before it wait, holding only cells already counted against
 *   `maxCellsPerSheet`). ROW-only rows are emitted at their place in that
 *   order, without waiting for any record. Hidden rows are emitted with
 *   `hidden: true`, never skipped.
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
import { absurd, Cause, Effect, Stream } from "effect"
import type * as Arr from "effect/Array"
import type * as Pull from "effect/Pull"
import * as Cell from "../../Cell.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row, SheetDetails } from "../../Workbook.ts"
import { readRichString, RecordCursor } from "../biff/cursor.ts"
import { type BiffRecord, CONTINUE, Embedded, EOF, readBof, SubstreamReader } from "../biff/records.ts"
import { decodeRk, ERROR_CODES } from "../binary.ts"
import { FixedReader, type StringFailure } from "../binary/layout.ts"
import { WorkMeter } from "../meter.ts"
import { numberCell } from "../numFmt.ts"
import { LayoutBuilder } from "../sheet/layout.ts"
import { HeldRows, NO_ROWS } from "../sheet/rows.ts"
import { yieldNow } from "../yield.ts"
import { readString } from "./fields.ts"
import type { Globals } from "./globals.ts"

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
 * theirs, or for an earlier row): at most 16 batches. Rows
 * stream in row order, so a row whose cells come both early and late in
 * the sheet holds back every row after it; real files split rows across a
 * few rows only (the corpus holds at most a few thousand cells). Past it,
 * the sheet is MalformedFile: MS-XLS puts a row's cells in its row block.
 * Tests may lower it (`SheetContext.heldCells`).
 */
export const HELD_CELLS = 16 * BATCH_CELLS

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

/*
 * The declared minimum lengths of the fixed fields each record is read
 * with (MS-XLS 2.4): a record is checked once against its minimum, then its
 * fields are read without a check each. Strings and cell lists after them
 * check their own bounds.
 */
/** Every cell and ROW record starts with its row: rw u16. */
const ROW_FIELD_MIN = 2
/** Cell: rw u16, col u16, ixfe u16. BLANK is just that; LABEL and RSTRING add a string. */
const CELL_MIN = 6
/** ROW: rw u16 ... grbit u16 at 12 (the flags with fDyZero). */
const ROW_MIN = 14
/** LABELSST: cell, isst u32. */
const LABELSST_MIN = 10
/** NUMBER: cell, num f64. FORMULA: cell, result (8 bytes; its last u16 0xFFFF for a special result). */
const NUMBER_MIN = 14
const FORMULA_MIN = 14
/** RK: cell, rk u32. */
const RK_MIN = 10
/** BOOLERR: cell, bBoolErr u8, fError u8. */
const BOOLERR_MIN = 8
/** MULRK, MULBLANK: rw u16, colFirst u16, the cells (6 or 2 bytes each), colLast u16. */
const MUL_MIN = 6
/** DIMENSIONS: rows (u32 each in BIFF8, u16 in BIFF5), then columns (u16 each). */
const DIMENSIONS_MIN_BIFF8 = 12
const DIMENSIONS_MIN_BIFF5 = 8
/** COLINFO: colFirst u16, colLast u16, coldx u16, ixfe u16, flags u16. */
const COLINFO_MIN = 10
/** MERGECELLS: cmcs u16, then 8 bytes per range (rwFirst, rwLast, colFirst, colLast: u16 each). */
const MERGECELLS_COUNT = 2
const MERGE_SIZE = 8

export interface SheetContext {
  readonly stream: Uint8Array
  readonly offset: number
  /** End of the substream: reading stops there (see `SheetEntry.end`). */
  readonly end: number
  readonly part: string
  readonly globals: Globals
  readonly limits: Limits
  /** The cap on held cells, `HELD_CELLS` by default (tests lower it). */
  readonly heldCells?: number
}

const malformed = (part: string, reason: string, offset?: number) =>
  new MalformedFile({ format: "xls", part, reason, ...(offset === undefined ? {} : { offset }) })

/**
 * Unknown error codes (BOOLERR, FORMULA results) give `#N/A`, not a failure:
 * as the XLSX reader, which keeps whatever error text a cell has.
 */
const errorCode = (code: number): string => ERROR_CODES.get(code) ?? "#N/A"

const truncated = (part: string, record: BiffRecord) => malformed(part, `truncated record 0x${record.type.toString(16)}`, record.offset)

/** The error of a failed string read in `record`: truncated, or LimitExceeded with the string's length. */
const stringFailed = (read: StringFailure, part: string, record: BiffRecord, limits: Limits): SpreadsheetError =>
  read._tag === "TooLong"
    ? new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual: read.count })
    : truncated(part, record)

/**
 * Opens the substream: checks the offset and the worksheet BOF; the reader
 * of its records after it (embedded substreams skipped, records of type
 * `continued` read with their CONTINUEs), or why not.
 */
const openSubstream = (context: SheetContext, continued?: number): SubstreamReader | SpreadsheetError => {
  if (context.offset < 0 || context.offset + 4 > context.stream.length) {
    return malformed(context.part, `sheet offset ${context.offset} is outside the workbook stream`)
  }
  const reader = new SubstreamReader(context.stream, context.part, context.offset, context.end, continued)
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
 * A first pass over the substream's records (headers, row numbers and ROW
 * flags only): the offset of each row's last cell record, and the rows ROW
 * records declare and hide. Some writers emit a row's cells in several
 * places, or ROW records after the cells; knowing all this up front lets the
 * main pass merge a row's cells, with its hidden flag, and still emit every
 * row, in row order, as soon as its cells are complete. A row declared only
 * by ROW records has no cell to wait for: its entry is -1, due at once. One
 * entry per row with cells or a ROW record (at most 65 536: rows are u16).
 */
class RowScan {
  readonly lastAt = new Map<number, number>()
  /** Rows with a ROW record. */
  readonly declared = new Set<number>()
  /** Rows with a ROW record that hides them (fDyZero). */
  readonly hidden = new Set<number>()
  done = false
  private readonly reader: SubstreamReader
  private readonly fields = new FixedReader()

  constructor(context: SheetContext, offset: number) {
    this.reader = new SubstreamReader(context.stream, context.part, offset, context.end)
  }

  /**
   * Reads up to `max` records; `done` at the sheet's EOF or where the records
   * stop (the main pass reports why). Returns its steps, for the work meter.
   */
  step(max: number): number {
    // progress: k up to max; every read() advances the reader
    for (let k = 0; k < max; k++) {
      const record = this.reader.read()
      if (record === undefined || record instanceof MalformedFile) {
        this.done = true
        return k + 1
      }
      if (record instanceof Embedded) continue
      if (record.type === EOF) {
        this.done = true
        return k + 1
      }
      // a cell or ROW record too short for its row is reported by the
      // main pass; any other record isn't read here
      const f = this.fields.fixed(record.data, ROW_FIELD_MIN)
      if (f === undefined) continue
      const row = f.u16(0)
      if (CELL_RECORDS.has(record.type)) this.lastAt.set(row, record.offset)
      else if (record.type === ROW) {
        this.declared.add(row)
        // the flags byte, if the record reaches it (a ROW cut before
        // ROW_MIN fails in the main pass, after rows before it are out)
        if (record.data.length > 12 && (f.u8(12) & 0x20) !== 0) this.hidden.add(row)
        if (!this.lastAt.has(row)) this.lastAt.set(row, -1)
      }
    }
    return max
  }
}

/**
 * The wait for a formula's text result: none, or after a FORMULA whose
 * string result comes in the next STRING record (MS-XLS: FORMULA, then
 * maybe SHRFMLA, ARRAY or TABLE, then STRING, any of them maybe continued
 * by CONTINUE records).
 */
export type FormulaWait =
  | { readonly _tag: "Idle" }
  | { readonly _tag: "AwaitingString"; readonly row: number; readonly column: number }

const IDLE: FormulaWait = { _tag: "Idle" }

/**
 * Builds rows from cell records and the `RowScan`'s ROW records, through
 * the shared `HeldRows`. Cells of a row are gathered until the record
 * holding its last cell is handled, then rows are emitted in row order.
 * Rows waiting for later cells hold only cells already counted against
 * `maxCellsPerSheet`.
 */
export class Rows {
  /** Rows ready to emit; `out[outAt]` is the next one (see `take`). */
  private out: ReadonlyArray<Row> = NO_ROWS
  private outAt = 0
  /** The rows' cells, limits and compaction. */
  private readonly rows: HeldRows
  /** Offset of each row's last cell record, -1 for a row with ROW records only (see `RowScan`). */
  private lastAt: ReadonlyMap<number, number> = new Map()
  /** Rows the scan placed, ascending; `order[nextRow]` is the next one to emit. */
  private order: Uint32Array = new Uint32Array(0)
  private nextRow = 0
  /** `lastAt` of `order[nextRow]` (Infinity when none is left). */
  private nextAt = Infinity
  /** Rows declared by ROW records, and those they hide (from the `RowScan`). */
  private declared: ReadonlySet<number> = new Set()
  private hidden: ReadonlySet<number> = new Set()
  /** Rows the scan didn't place, in row order (set by `end`); `left[leftAt]` is the next one. */
  private left: Uint32Array | undefined
  private leftAt = 0
  /** Emitting rows stops once `effort` reaches it, and goes on in the next batch. */
  budget = Infinity
  /**
   * Steps of its loops (cells, rows, queue steps), and of the caller's
   * record loop, since the work meter was last told (see `takeSteps`).
   */
  steps = 0
  /** The rows' sorting work already handed to the work meter. */
  private sortedTold = 0
  /** Whether a FORMULA waits for its text result in the next STRING record (see `handle`). */
  wait: FormulaWait = IDLE
  /** Pointed at each record whose fixed fields are read (see `handle`). */
  readonly fields = new FixedReader()

  constructor(context: SheetContext) {
    this.rows = new HeldRows(context.limits)
  }

  /** Cells in rows not emitted yet (see `HELD_CELLS`). */
  get held(): number {
    return this.rows.held
  }

  /** Cells read, cells sorted (compacting a row) and rows emitted: the work that a batch spends, up to `budget`. */
  get effort(): number {
    return this.rows.work
  }

  /** The steps since the last call, sorting steps included: the work meter's units. */
  takeSteps(): number {
    const sorted = this.rows.sorted
    const steps = this.steps + sorted - this.sortedTold
    this.sortedTold = sorted
    this.steps = 0
    return steps
  }

  /** Whether row `index` waits for its formula's STRING. */
  private waiting(index: number): boolean {
    const wait = this.wait
    return wait._tag === "AwaitingString" && wait.row === index
  }

  /** Where each row's cells end and the rows ROW records declare and hide, from the `RowScan`. */
  start(lastAt: ReadonlyMap<number, number>, declared: ReadonlySet<number> = new Set(), hidden: ReadonlySet<number> = new Set()): void {
    this.lastAt = lastAt
    this.declared = declared
    this.hidden = hidden
    this.order = Uint32Array.from(lastAt.keys()).sort()
    this.nextRow = 0
    this.nextAt = this.order.length > 0 ? lastAt.get(this.order[0]!)! : Infinity
  }

  /** Emits row `index` with its cells, or without any if only ROW records declared it. */
  private finishRow(index: number): SpreadsheetError | undefined {
    const hidden = this.hidden.has(index)
    const finished = this.rows.finishRow(index, hidden)
    if (finished === true) return undefined
    if (finished !== false) return finished
    return this.declared.has(index) ? this.rows.finishEmptyRow(index, hidden) : undefined
  }

  /** Whether rows wait to be taken. */
  queued(): boolean {
    return this.outAt < this.out.length || this.rows.ready
  }

  /** The next chunk of rows: at least one, at most `BATCH_RECORDS` rows or until `BATCH_CELLS` cells are passed. */
  take(): Arr.NonEmptyReadonlyArray<Row> {
    if (this.outAt === this.out.length) {
      this.out = this.rows.take()
      this.outAt = 0
    }
    const from = this.outAt
    let cells = 0
    // progress: outAt increases up to out.length
    do {
      this.steps++
      cells += this.out[this.outAt++]!.cells.length
    } while (this.outAt < this.out.length && this.outAt - from < BATCH_RECORDS && cells < BATCH_CELLS)
    const chunk = this.out.slice(from, this.outAt) as unknown as Arr.NonEmptyReadonlyArray<Row>
    if (this.outAt === this.out.length) {
      this.out = NO_ROWS
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
      this.steps++
      const index = this.order[this.nextRow]!
      if (this.waiting(index)) return undefined
      this.nextRow++
      this.nextAt = this.nextRow < this.order.length ? this.lastAt.get(this.order[this.nextRow]!)! : Infinity
      const error = this.finishRow(index)
      if (error !== undefined) return error
    }
    return undefined
  }

  /** Whether rows complete at `offset` are left for `complete` (it stopped at the budget). */
  behind(offset: number): boolean {
    return this.nextRow < this.order.length && this.nextAt <= offset && !this.waiting(this.order[this.nextRow]!)
  }

  /** Adds `cell` to row `row`. */
  cell(row: number, cell: Cell.Cell): SpreadsheetError | undefined {
    this.steps++
    return this.rows.addCell(row, cell)
  }

  /**
   * After the EOF, once `complete(Infinity)` isn't `behind`: emits the
   * rest, until `effort` reaches `budget`; `true` once all is out (else
   * call it again in the next batch).
   */
  end(): SpreadsheetError | boolean {
    // rows the scan didn't place (none when both passes agree), in row order
    const left = this.left ??= this.rows.indexes()
    // progress: leftAt increases up to left.length
    while (this.leftAt < left.length) {
      if (this.effort >= this.budget) return false
      this.steps++
      const failed = this.finishRow(left[this.leftAt++]!)
      if (failed !== undefined) return failed
    }
    return true
  }
}

/** A number cell, a Date when its XF has a date format. */
const xfNumber = (globals: Globals, column: number, xf: number, value: number): Cell.Cell =>
  numberCell(column, value, globals.styles.format(xf), globals.date1904)

/**
 * Handles one record of the substream; `true` at its EOF. An embedded
 * substream (a chart object) is skipped; its BOF ends a formula's wait.
 */
const handle = (record: BiffRecord | Embedded, rows: Rows, context: SheetContext): SpreadsheetError | boolean => {
  const { globals, part } = context
  // a formula's string result comes right after it, or after its SHRFMLA, ARRAY or TABLE
  // (any of them maybe continued by CONTINUE records, which are part of it). Any other
  // record ends the wait, an embedded substream's BOF too (no FORMULA is read inside one,
  // so its other records come with no wait). FORMULA starts the wait and STRING takes it
  // (see their cases)
  const wait = rows.wait
  switch (wait._tag) {
    case "Idle":
      break
    case "AwaitingString": {
      if (!(record instanceof Embedded)) {
        const type = record.type
        if (type === STRING || type === SHRFMLA || type === ARRAY || type === TABLE || type === CONTINUE) break
      }
      rows.wait = IDLE
      const error = rows.cell(wait.row, Cell.text(wait.column, ""))
      if (error !== undefined) return error
      break
    }
    default:
      return absurd(wait)
  }
  if (record instanceof Embedded) return false
  const { data } = record
  const fields = rows.fields
  switch (record.type) {
    case EOF:
      return true
    case ROW:
      // the `RowScan` already took the declared row and its hidden flag
      return fields.fixed(data, ROW_MIN) === undefined ? truncated(part, record) : false
    case LABELSST: {
      const f = fields.fixed(data, LABELSST_MIN)
      if (f === undefined) return truncated(part, record)
      const isst = f.u32(6)
      const value = globals.strings[isst]
      if (value === undefined) return malformed(part, `shared string ${isst} doesn't exist`, record.offset)
      return rows.cell(f.u16(0), Cell.text(f.u16(2), value)) ?? false
    }
    case LABEL:
    case RSTRING: {
      const f = fields.fixed(data, CELL_MIN)
      if (f === undefined) return truncated(part, record)
      const row = f.u16(0)
      const column = f.u16(2)
      const text = readString(data, CELL_MIN, globals.text, 2)
      if (text._tag !== "Ok") return stringFailed(text, part, record, context.limits)
      return rows.cell(row, Cell.text(column, text.value)) ?? false
    }
    case NUMBER: {
      const f = fields.fixed(data, NUMBER_MIN)
      if (f === undefined) return truncated(part, record)
      return rows.cell(f.u16(0), xfNumber(globals, f.u16(2), f.u16(4), f.f64(6))) ?? false
    }
    case RK: {
      const f = fields.fixed(data, RK_MIN)
      if (f === undefined) return truncated(part, record)
      return rows.cell(f.u16(0), xfNumber(globals, f.u16(2), f.u16(4), decodeRk(f.u32(6)))) ?? false
    }
    case MULRK:
    case MULBLANK: {
      const f = fields.fixed(data, MUL_MIN)
      if (f === undefined) return truncated(part, record)
      const row = f.u16(0)
      const first = f.u16(2)
      const last = f.u16(data.length - 2)
      const width = record.type === MULRK ? 6 : 2
      const count = (data.length - MUL_MIN) / width
      if (!Number.isInteger(count) || last - first + 1 !== count) {
        return malformed(part, `record 0x${record.type.toString(16)} doesn't match its column range`, record.offset)
      }
      // the cells fill the record between colFirst and colLast: cell k's
      // ixfe at 4 + width * k (and its rk after it) ends by data.length - 2.
      // progress: k up to count (bounded by the record length)
      for (let k = 0; k < count; k++) {
        const xf = f.u16(4 + width * k)
        const cell = record.type === MULRK
          ? xfNumber(globals, first + k, xf, decodeRk(f.u32(6 + width * k)))
          : Cell.empty(first + k)
        const error = rows.cell(row, cell)
        if (error !== undefined) return error
      }
      return false
    }
    case BOOLERR: {
      const f = fields.fixed(data, BOOLERR_MIN)
      if (f === undefined) return truncated(part, record)
      const row = f.u16(0)
      const column = f.u16(2)
      const value = f.u8(6)
      if (f.u8(7) === 1) return rows.cell(row, Cell.error(column, errorCode(value))) ?? false
      return rows.cell(row, Cell.boolean(column, value !== 0)) ?? false
    }
    case BLANK: {
      const f = fields.fixed(data, CELL_MIN)
      if (f === undefined) return truncated(part, record)
      return rows.cell(f.u16(0), Cell.empty(f.u16(2))) ?? false
    }
    case FORMULA:
    case FORMULA_BIFF4: {
      const f = fields.fixed(data, FORMULA_MIN)
      if (f === undefined) return truncated(part, record)
      const row = f.u16(0)
      const column = f.u16(2)
      if (f.u16(12) === 0xFFFF) {
        // a special result: string (in the next STRING record), boolean, error, empty string
        const kind = f.u8(6)
        const value = f.u8(8)
        if (kind === 0) {
          rows.wait = { _tag: "AwaitingString", row, column }
          return false
        }
        if (kind === 1) return rows.cell(row, Cell.boolean(column, value !== 0)) ?? false
        if (kind === 2) return rows.cell(row, Cell.error(column, errorCode(value))) ?? false
        return rows.cell(row, Cell.text(column, "")) ?? false
      }
      return rows.cell(row, xfNumber(globals, column, f.u16(4), f.f64(6))) ?? false
    }
    case STRING: {
      const pending = rows.wait
      if (pending._tag === "Idle") return false // a STRING without a formula: ignored
      rows.wait = IDLE
      // BIFF8: an XLUnicodeString, maybe continued (each CONTINUE restarts
      // with an option byte; the reader checked them and moved past them);
      // its character count is checked before any character is read.
      // BIFF5: a u16-count code-page string, never continued
      const read = globals.biff8
        ? readRichString(new RecordCursor(context.stream, record, context.end), context.limits.maxStringLength)
        : readString(data, 0, globals.text, 2)
      if (read._tag !== "Ok") return stringFailed(read, part, record, context.limits)
      return rows.cell(pending.row, Cell.text(pending.column, read.value)) ?? false
    }
    default:
      return false
  }
}

/**
 * Where a rows reader is: its substream didn't open; its record headers
 * are being scanned (`RowScan`); its records are being read (rows complete
 * up to `handledAt`, the offset of the last record handled, may wait for
 * budget); its EOF was handled and only rows are left to emit; or all is out.
 */
type Phase =
  | { readonly _tag: "Failed"; readonly error: SpreadsheetError }
  | { readonly _tag: "Scanning"; readonly reader: SubstreamReader; readonly scan: RowScan }
  | { readonly _tag: "Reading"; readonly reader: SubstreamReader; handledAt: number }
  | { readonly _tag: "Draining" }
  | { readonly _tag: "Done" }

const DRAINING: Phase = { _tag: "Draining" }
const DONE: Phase = { _tag: "Done" }

/** The rows of a worksheet substream. Re-runnable. */
export const sheetRows = (context: SheetContext): Stream.Stream<Row, SpreadsheetError> =>
  Stream.fromPull(Effect.map(WorkMeter, (meter) => {
    // a BIFF8 STRING comes with its CONTINUEs (a BIFF5 one has none)
    const opened = openSubstream(context, context.globals.biff8 ? STRING : undefined)
    let phase: Phase = opened instanceof SubstreamReader
      ? { _tag: "Scanning", reader: opened, scan: new RowScan(context, opened.offset) }
      : { _tag: "Failed", error: opened }
    const rows = new Rows(context)
    const heldCells = context.heldCells ?? HELD_CELLS
    /** Hands the steps of the last batch to the meter. */
    const report = Effect.sync(() => {
      meter.add("steps", rows.takeSteps())
    })
    const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, SpreadsheetError> = Effect.gen(function*() {
      // progress: each iteration takes queued rows, or scans up to
      // BATCH_RECORDS record headers, or reads up to BATCH_RECORDS records
      // (the scan's or the reader's offset grows with each), or emits rows
      // (at least one, or the phase moves on), or ends the stream
      while (true) {
        yield* report
        yield* yieldNow
        if (rows.queued()) return rows.take()
        const current: Phase = phase
        switch (current._tag) {
          case "Failed":
            return yield* Effect.fail(current.error)
          case "Scanning": {
            const { scan } = current
            rows.steps += scan.step(BATCH_RECORDS)
            if (scan.done) {
              rows.start(scan.lastAt, scan.declared, scan.hidden)
              phase = { _tag: "Reading", reader: current.reader, handledAt: -1 }
            }
            continue
          }
          case "Reading": {
            const { reader } = current
            rows.budget = rows.effort + BATCH_CELLS
            // rows a budget left complete but not emitted come first: each
            // call emits at least one, or none is left
            const caught = rows.complete(current.handledAt)
            if (caught !== undefined) return yield* Effect.fail(caught)
            // progress: k up to BATCH_RECORDS; every read() advances the reader
            for (let k = 0; k < BATCH_RECORDS && rows.effort < rows.budget && !rows.behind(current.handledAt); k++) {
              rows.steps++
              const record = reader.read()
              if (record === undefined) return yield* Effect.fail(malformed(context.part, "the sheet has no EOF"))
              if (record instanceof MalformedFile) return yield* Effect.fail(record)
              const result = handle(record, rows, context)
              if (result === true) {
                phase = DRAINING
                break
              }
              if (result !== false) return yield* Effect.fail(result)
              if (rows.held > heldCells) {
                return yield* Effect.fail(malformed(context.part, `rows split too far apart: over ${heldCells} cells held for rows not complete`, record.offset))
              }
              current.handledAt = record.offset
              const error = rows.complete(record.offset)
              if (error !== undefined) return yield* Effect.fail(error)
            }
            if (rows.queued()) return rows.take()
            continue
          }
          case "Draining": {
            rows.budget = rows.effort + BATCH_CELLS
            // rows a budget left complete but not emitted, then the rest: each
            // call emits at least one, or none is left
            const caught = rows.complete(Infinity)
            if (caught !== undefined) return yield* Effect.fail(caught)
            if (!rows.behind(Infinity)) {
              const result = rows.end()
              if (result === true) phase = DONE
              else if (result !== false) return yield* Effect.fail(result)
            }
            if (rows.queued()) return rows.take()
            continue
          }
          case "Done":
            return yield* Cause.done()
          default:
            return absurd(current)
        }
      }
    })
    // the steps of the last batch count also when the pull fails
    return Effect.ensuring(pull, report)
  }))

/**
 * Hidden columns (COLINFO), merges (MERGECELLS) and the declared dimension
 * (DIMENSIONS) of a worksheet substream, in one pass over its records
 * (yielding every `BATCH_RECORDS` records).
 */
export const sheetDetails = (context: SheetContext): Effect.Effect<SheetDetails, SpreadsheetError> =>
  Effect.flatMap(WorkMeter, (meter) => {
    const walked: Walked = { steps: 0 }
    const report = Effect.sync(() => {
      meter.add("steps", walked.steps)
      walked.steps = 0
    })
    return Effect.ensuring(walkDetails(context, walked, report), report)
  })

/** Records walked since the work meter was last told. */
interface Walked {
  steps: number
}

/** `sheetDetails`' walk: it counts each record in `walked`, and `report`s the count before each yield. */
const walkDetails = (context: SheetContext, walked: Walked, report: Effect.Effect<void>): Effect.Effect<SheetDetails, SpreadsheetError> =>
  Effect.gen(function*() {
    const opened = openSubstream(context)
    if (!(opened instanceof SubstreamReader)) return yield* Effect.fail(opened)
    const { part, limits, globals } = context
    const layout = new LayoutBuilder(limits)
    const fields = new FixedReader()
    // progress: every read() advances the reader
    for (let k = 1;; k++) {
      walked.steps++
      if (k % BATCH_RECORDS === 0) {
        yield* report
        yield* yieldNow
      }
      const record = opened.read()
      if (record === undefined) return yield* Effect.fail(malformed(part, "the sheet has no EOF"))
      if (record instanceof MalformedFile) return yield* Effect.fail(record)
      if (record instanceof Embedded) continue
      if (record.type === EOF) break
      const { data } = record
      if (record.type === DIMENSIONS) {
        // a cut DIMENSIONS is ignored, as is an empty one
        const biff8 = globals.biff8
        const f = fields.fixed(data, biff8 ? DIMENSIONS_MIN_BIFF8 : DIMENSIONS_MIN_BIFF5)
        if (f === undefined) continue
        const [r1, r2, c1, c2] = biff8 ? [f.u32(0), f.u32(4), f.u16(8), f.u16(10)] : [f.u16(0), f.u16(2), f.u16(4), f.u16(6)]
        if (r2 > r1 && c2 > c1) layout.dimension({ firstRow: r1, firstColumn: c1, lastRow: r2 - 1, lastColumn: c2 - 1 })
      } else if (record.type === COLINFO) {
        const tooMany = layout.columnRecord()
        if (tooMany !== undefined) return yield* Effect.fail(tooMany)
        const f = fields.fixed(data, COLINFO_MIN)
        if (f === undefined) return yield* Effect.fail(malformed(part, "truncated COLINFO", record.offset))
        // the part at or past maxColumns is dropped, as in XLSX
        if ((f.u16(8) & 0x01) !== 0) layout.hideColumns(f.u16(0), f.u16(2))
      } else if (record.type === MERGECELLS) {
        const f = fields.fixed(data, MERGECELLS_COUNT)
        if (f === undefined || MERGECELLS_COUNT + MERGE_SIZE * f.u16(0) > data.length) {
          return yield* Effect.fail(malformed(part, "truncated MERGECELLS", record.offset))
        }
        const count = f.u16(0)
        // progress: j up to count (the ranges fit in the record: checked above)
        for (let j = 0; j < count; j++) {
          const at = MERGECELLS_COUNT + MERGE_SIZE * j
          const [r1, r2, c1, c2] = [f.u16(at), f.u16(at + 2), f.u16(at + 4), f.u16(at + 6)]
          const tooMany = layout.merge({ firstRow: Math.min(r1, r2), firstColumn: Math.min(c1, c2), lastRow: Math.max(r1, r2), lastColumn: Math.max(c1, c2) })
          if (tooMany !== undefined) return yield* Effect.fail(tooMany)
        }
      }
    }
    return layout.result()
  })

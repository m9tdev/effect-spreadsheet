/**
 * A worksheet substream of a BIFF8/BIFF5 workbook stream -> rows, streamed.
 *
 * The substream starts at its BOUNDSHEET offset with a worksheet BOF and
 * runs to its EOF. Its rows are read in two steps, both in bounded turns
 * with a yield to the event loop between them (`sheetRows`):
 *
 * 1. **The index** (`indexRecords`): one pass over the substream's records,
 *    in stream order. It checks every record as it will be read (record
 *    lengths, strings, shared-string indexes, column ranges, and the
 *    `maxColumns` and `maxCellsPerSheet` limits, which count cells in stream
 *    order), and notes where each row's cell records are: one range per row
 *    for the common case, where a row's cells come together, and the
 *    offsets of the records that come later, out of place (a row whose
 *    cells some writers put in several places), chained by row. ROW records
 *    give hidden rows (fDyZero) and rows without cells, wherever they come.
 *    It holds a few numbers per row and one per out-of-place record, never
 *    a cell. The first failure is kept with the offset where it happens;
 *    the pass then goes on over the record headers only, to the EOF.
 * 2. **The rows** (`emitRows`): in row order, each row's records are read
 *    again from the index and decoded into its cells, merged (the last cell
 *    of a column wins, in stream order), and the row goes out. A row goes
 *    out only if it was complete before the failure, as when the records
 *    were read in stream order: its records, and those of the rows before
 *    it, all come before the failure's offset (a FORMULA with a string
 *    result is complete at its STRING). Then the failure is reported. ROW-only
 *    rows are complete from the start. Hidden rows are emitted with
 *    `hidden: true`, never skipped.
 *
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
import type { Row, SheetDetails } from "../../Workbook.ts"
import { readRichString, RecordCursor } from "../biff/cursor.ts"
import { type BiffRecord, CONTINUE, Embedded, EOF, readBof, SubstreamReader } from "../biff/records.ts"
import { decodeRk, ERROR_CODES } from "../binary.ts"
import { FixedReader, type StringFailure } from "../binary/layout.ts"
import { WorkMeter } from "../meter.ts"
import { numberCell } from "../numFmt.ts"
import { LayoutBuilder } from "../sheet/layout.ts"
import { RowCells } from "../sheet/rows.ts"
import { Turn } from "../walk.ts"
import { yieldNow } from "../yield.ts"
import { readString, stringFailure } from "./fields.ts"
import type { Globals } from "./globals.ts"

/** Records per turn: bounded work between event-loop turns. */
export const BATCH_RECORDS = 4096
/**
 * Cells per turn, read or emitted: a turn ends once it passes this (one
 * MULBLANK holds up to 4 109 cells), and so does a chunk of rows (a row
 * holds at most `maxColumns` cells, see `compact`).
 */
export const BATCH_CELLS = 65_536
/** What a record costs in a turn, against a cell's 1: a turn holds `BATCH_RECORDS` records or `BATCH_CELLS` cells. */
const RECORD_COST = BATCH_CELLS / BATCH_RECORDS
/**
 * How far past its budget a turn may go to finish the row it is in, so a
 * row's cells rarely wait across a yield: cells kept across event-loop
 * turns survive young-generation collections and get promoted, which made
 * wide rows (16 384 cells) split over turns read much slower.
 */
const ROW_SLACK = BATCH_CELLS / 4

/**
 * Records of a sheet out of place: cell records of a row that come after
 * another row's cells started (a row split far apart, or rows out of
 * order). Each costs one offset in the index. Real files split rows across
 * a few rows only; past this many, the sheet is MalformedFile (MS-XLS puts
 * a row's cells in its row block). Tests may lower it
 * (`SheetContext.displacedRecords`).
 */
export const DISPLACED_RECORDS = 1_048_576

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
  /** End of the substream: reading stops there (see `SheetStart`). */
  readonly end: number
  readonly part: string
  readonly globals: Globals
  readonly limits: Limits
  /** The cap on records out of place, `DISPLACED_RECORDS` by default (tests lower it). */
  readonly displacedRecords?: number
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
 * Opens the substream at its start (an offset the globals checked, see
 * `SheetStart`): checks its BOF; the reader of its records after it
 * (embedded substreams skipped, records of type `continued` read with
 * their CONTINUEs), or why not.
 */
const openSubstream = (context: SheetContext, continued?: number): SubstreamReader | SpreadsheetError => {
  const reader = new SubstreamReader(context.stream, context.part, context.offset, context.end, continued)
  const bof = reader.next()
  if (bof === undefined) return malformed(context.part, "the sheet has no BOF", context.offset)
  if (bof instanceof MalformedFile) return bof
  const parsed = readBof(bof)
  if (parsed === undefined) return malformed(context.part, "the sheet doesn't start with a BOF", context.offset)
  return reader
}

/** Whether a record of `type` is a cell record: each starts with its row (u16). */
const isCell = (type: number): boolean =>
  type === LABELSST || type === LABEL || type === RSTRING || type === NUMBER || type === RK || type === MULRK ||
  type === BOOLERR || type === BLANK || type === MULBLANK || type === FORMULA || type === FORMULA_BIFF4

/**
 * Whether a record of `type` keeps a FORMULA's wait for its string result:
 * its STRING comes right after it, or after its SHRFMLA, ARRAY or TABLE
 * (any of them maybe continued by CONTINUE records, which are part of it).
 * Any other record ends the wait (the result is ""), an embedded
 * substream's BOF too.
 */
const keepsWait = (type: number): boolean => type === STRING || type === SHRFMLA || type === ARRAY || type === TABLE || type === CONTINUE

/** A number cell, a Date when its XF has a date format. */
const xfNumber = (globals: Globals, column: number, xf: number, value: number): Cell.Cell =>
  numberCell(column, value, globals.styles.format(xf), globals.date1904)

/** Bits of `SheetIndex.flags`. */
const DECLARED = 1
const HIDDEN = 2
/** The row has a range of cell records (`start` to `end`). */
const RANGED = 4

/** Where a failure stands when the sheet has no EOF: after every record. */
const AFTER_ALL = Number.MAX_SAFE_INTEGER

/** A growable array of u32s. */
class U32List {
  items = new Uint32Array(64)
  length = 0

  push(value: number): void {
    if (this.length === this.items.length) {
      const grown = new Uint32Array(this.items.length * 2)
      grown.set(this.items)
      this.items = grown
    }
    this.items[this.length++] = value
  }
}

/**
 * Numbers noted per slot, each slot's in the order noted: one chain per
 * slot through `next` (entries are numbered from 1; 0 ends a chain), so a
 * slot's numbers are read in order without grouping them first.
 */
class Chains {
  /** The numbers, and each one's next in its slot's chain (0 for none). */
  readonly values = new U32List()
  readonly next = new U32List()
  /** Each slot's first and last entry (0 for none). */
  head = new Uint32Array(64)
  private tail = new Uint32Array(64)

  /** Notes `value` for `slot`, after the slot's earlier ones. */
  push(slot: number, value: number): void {
    this.values.push(value)
    this.next.push(0)
    const entry = this.values.length
    const last = this.tail[slot]!
    if (last === 0) this.head[slot] = entry
    else this.next.items[last - 1] = entry
    this.tail[slot] = entry
  }

  /** Room for `size` slots. */
  grow(size: number): void {
    const head = new Uint32Array(size)
    head.set(this.head)
    this.head = head
    const tail = new Uint32Array(size)
    tail.set(this.tail)
    this.tail = tail
  }
}

/**
 * Where a sheet's rows are, by row: one slot per row with cell records or
 * a ROW record (at most 65 536: rows are u16), in the order they are first
 * seen. A slot holds its row, the range of its cell records that came
 * together first (`start` to `end`, from the first record's header to the
 * end of the last one), its flags, and `complete`: the offset after which
 * the row was complete in stream order (its last cell record, or for a
 * FORMULA waiting for its string result, the record that ends the wait;
 * -1 for a row with ROW records only). Cell records of a row that come
 * after another row's (out of place) are kept apart, by offset, in the
 * slot's chain.
 */
class SheetIndex {
  slots = 0
  row = new Uint32Array(64)
  start = new Uint32Array(64)
  end = new Uint32Array(64)
  complete = new Float64Array(64)
  flags = new Uint8Array(64)
  private readonly bySlot = new Map<number, number>()
  /** The last row looked up and its slot: a row's records mostly come together. */
  private lastRow = -1
  private lastSlot = -1
  /** Out-of-place records: their offsets, by slot, in stream order. */
  readonly displaced = new Chains()
  /**
   * FORMULAs whose string result doesn't come in the record right after
   * them (a SHRFMLA, ARRAY, TABLE or CONTINUE comes first): the offset of
   * their STRING (0 for none: the result is ""), by slot, in stream order.
   * A row's FORMULAs are decoded in that order, and take them in turn.
   */
  readonly noted = new Chains()

  /** The slot of row `row`, made if new (its `complete` -1 until a cell record of it comes). */
  slotOf(row: number): number {
    if (row === this.lastRow) return this.lastSlot
    let slot = this.bySlot.get(row)
    if (slot === undefined) {
      slot = this.slots++
      if (slot === this.row.length) this.grow()
      this.row[slot] = row
      this.complete[slot] = -1
      this.bySlot.set(row, slot)
    }
    this.lastRow = row
    this.lastSlot = slot
    return slot
  }

  private grow(): void {
    const size = this.row.length * 2
    const row = new Uint32Array(size)
    row.set(this.row)
    this.row = row
    const start = new Uint32Array(size)
    start.set(this.start)
    this.start = start
    const end = new Uint32Array(size)
    end.set(this.end)
    this.end = end
    const complete = new Float64Array(size)
    complete.set(this.complete)
    this.complete = complete
    const flags = new Uint8Array(size)
    flags.set(this.flags)
    this.flags = flags
    this.displaced.grow(size)
    this.noted.grow(size)
  }

  /** The slots in row order. */
  ordered(): Uint32Array {
    // a row is a u16 and a slot below 65 536: row x 65 536 + slot sorts by row
    const keys = new Uint32Array(this.slots)
    // progress: k up to slots
    for (let k = 0; k < this.slots; k++) keys[k] = this.row[k]! * 0x10000 + k
    keys.sort()
    // progress: k up to slots
    for (let k = 0; k < this.slots; k++) keys[k] = keys[k]! & 0xFFFF
    return keys
  }

}

/**
 * What the rows walk hands to its driver: the work done since the work
 * meter was last told (steps of every loop: records, cells, sorting, rows),
 * and the rows ready to go out. The rows are handed over here rather than
 * yielded: a suspended generator may keep the value it yielded last, and a
 * chunk of rows kept one turn longer than needed cost a wide sheet its
 * cells' young-generation collections.
 */
interface Tally {
  steps: number
  rows: Array<Row>
}

/** The index pass's outcome: the index, and the first failure in stream order with where it stands. */
interface Indexed {
  readonly index: SheetIndex
  readonly failure: SpreadsheetError | undefined
  /** Offset of the record that failed (`AFTER_ALL` for a missing EOF); Infinity without a failure. */
  readonly failAt: number
}

/**
 * The index pass (see the module's comment): reads the substream's records
 * from `reader` in stream order up to its EOF, checking each as `emitRows`
 * will read it, and builds the index. Yields when the turn is spent.
 */
function* indexRecords(context: SheetContext, reader: SubstreamReader, turn: Turn, tally: Tally): Generator<void, Indexed, void> {
  const { globals, part, limits } = context
  const { maxColumns, maxCellsPerSheet } = limits
  const displacedCap = context.displacedRecords ?? DISPLACED_RECORDS
  const index = new SheetIndex()
  const fields = new FixedReader()
  let failure: SpreadsheetError | undefined
  let failAt = Infinity
  /** Cells read so far in stream order, repeats included (for maxCellsPerSheet). */
  let cells = 0
  /** The slot of the FORMULA waiting for its string result (-1 for none), its column, and the offset after it. */
  let waitSlot = -1
  let waitColumn = 0
  let waitNext = 0
  /** The slot whose cell records came last, and whether they are out of place (see `place`). */
  let runSlot = -1
  let runDisplaced = false

  const fail = (error: SpreadsheetError, at: number): void => {
    failure = error
    failAt = at
  }
  /**
   * Counts `count` cells in columns `first` on, as they are read: each one
   * fails at or past `maxColumns`, then past `maxCellsPerSheet` cells (the
   * column first, cell by cell). The error, if one of them fails.
   */
  const countCells = (first: number, count: number): SpreadsheetError | undefined => {
    if (count <= 0) return undefined
    // the first cell whose column is out, and the first one over the count
    const outColumn = Math.max(0, maxColumns - first)
    const overCount = Math.max(0, maxCellsPerSheet - cells)
    if (outColumn < count || overCount < count) {
      if (outColumn <= overCount) return new LimitExceeded({ limit: "maxColumns", max: maxColumns, actual: first + outColumn + 1 })
      return new LimitExceeded({ limit: "maxCellsPerSheet", max: maxCellsPerSheet, actual: maxCellsPerSheet + 1 })
    }
    cells += count
    return undefined
  }
  /** The wait for a formula's string result ends at the record at `at`: the row is complete there. */
  const waited = (at: number, string: boolean): void => {
    index.complete[waitSlot] = at
    if (at !== waitNext) {
      index.noted.push(waitSlot, string ? at : 0)
    }
    waitSlot = -1
  }
  /** Ends the wait for a formula's string result at the record at `at` without a STRING: its result is "" (one cell). */
  const endWait = (at: number): void => {
    waited(at, false)
    const error = countCells(waitColumn, 1)
    if (error !== undefined) fail(error, at)
  }
  /**
   * Notes the cell record `record` of slot `slot`: in the slot's range if
   * the slot's records come together so far, else out of place (a record
   * of a row whose range another row's records ended). Fails past the cap
   * on records out of place.
   */
  const place = (slot: number, record: BiffRecord): void => {
    const end = record.offset + 4 + record.data.length
    if (slot !== runSlot) {
      runSlot = slot
      runDisplaced = (index.flags[slot]! & RANGED) !== 0
      if (!runDisplaced) {
        index.flags[slot] = index.flags[slot]! | RANGED
        index.start[slot] = record.offset
        index.end[slot] = end
        return
      }
    } else if (!runDisplaced) {
      index.end[slot] = end
      return
    }
    if (index.displaced.values.length >= displacedCap) {
      fail(malformed(part, `rows split too far apart: over ${displacedCap} cell records out of their row's place`, record.offset), record.offset)
      return
    }
    index.displaced.push(slot, record.offset)
  }
  /**
   * Checks the cell record `record` of slot `slot` as `decode` reads it,
   * counts its cells, and places it; or fails (see `fail`). A FORMULA with
   * a string result waits for it instead (its cell is counted when the
   * wait ends).
   */
  const cellRecord = (record: BiffRecord, slot: number): void => {
    const { data } = record
    let error: SpreadsheetError | undefined
    switch (record.type) {
      case LABELSST: {
        const f = fields.fixed(data, LABELSST_MIN)
        if (f === undefined) error = truncated(part, record)
        else if (globals.strings[f.u32(6)] === undefined) error = malformed(part, `shared string ${f.u32(6)} doesn't exist`, record.offset)
        else error = countCells(f.u16(2), 1)
        break
      }
      case LABEL:
      case RSTRING: {
        const f = fields.fixed(data, CELL_MIN)
        if (f === undefined) error = truncated(part, record)
        else {
          const column = f.u16(2)
          const failed = stringFailure(data, CELL_MIN, globals.text, 2)
          error = failed !== undefined ? stringFailed(failed, part, record, limits) : countCells(column, 1)
        }
        break
      }
      case NUMBER:
      case RK:
      case BOOLERR:
      case BLANK: {
        const min = record.type === NUMBER ? NUMBER_MIN : record.type === RK ? RK_MIN : record.type === BOOLERR ? BOOLERR_MIN : CELL_MIN
        const f = fields.fixed(data, min)
        error = f === undefined ? truncated(part, record) : countCells(f.u16(2), 1)
        break
      }
      case MULRK:
      case MULBLANK: {
        const f = fields.fixed(data, MUL_MIN)
        if (f === undefined) error = truncated(part, record)
        else {
          const first = f.u16(2)
          const last = f.u16(data.length - 2)
          const count = (data.length - MUL_MIN) / (record.type === MULRK ? 6 : 2)
          error = !Number.isInteger(count) || last - first + 1 !== count
            ? malformed(part, `record 0x${record.type.toString(16)} doesn't match its column range`, record.offset)
            : countCells(first, count)
        }
        break
      }
      default: {
        // FORMULA, FORMULA_BIFF4
        const f = fields.fixed(data, FORMULA_MIN)
        if (f === undefined) error = truncated(part, record)
        else if (f.u16(12) === 0xFFFF && f.u8(6) === 0) {
          // a string result, in the STRING record that comes next
          waitSlot = slot
          waitColumn = f.u16(2)
          waitNext = record.offset + 4 + data.length
          index.complete[slot] = Infinity
        } else error = countCells(f.u16(2), 1)
      }
    }
    if (error !== undefined) fail(error, record.offset)
    else place(slot, record)
  }

  // progress: every read() advances the reader; the loop ends at the EOF or
  // where the records stop
  for (;;) {
    const record = reader.read()
    tally.steps++
    if (record === undefined) {
      if (failure === undefined) fail(malformed(part, "the sheet has no EOF"), AFTER_ALL)
      break
    }
    if (record instanceof MalformedFile) {
      if (failure === undefined) fail(record, record.offset ?? AFTER_ALL)
      break
    }
    if (record instanceof Embedded) {
      if (failure === undefined && waitSlot !== -1) endWait(record.offset)
    } else {
      const { type, data } = record
      if (failure === undefined && waitSlot !== -1 && !keepsWait(type)) endWait(record.offset)
      if (type === EOF) break
      if (isCell(type)) {
        // the row of every cell record long enough for it, as read in
        // stream order, also after a failure: rows not complete before it
        // don't go out
        const f = fields.fixed(data, ROW_FIELD_MIN)
        if (f !== undefined) {
          const slot = index.slotOf(f.u16(0))
          if (index.complete[slot]! < record.offset) index.complete[slot] = record.offset
          if (failure === undefined) cellRecord(record, slot)
        } else if (failure === undefined) fail(truncated(part, record), record.offset)
      } else if (type === ROW) {
        // the declared row and its hidden flag, wherever the ROW record is
        // (the flags byte if the record reaches it); a ROW cut before
        // ROW_MIN fails where it is
        const f = fields.fixed(data, ROW_FIELD_MIN)
        if (f !== undefined) {
          const slot = index.slotOf(f.u16(0))
          index.flags[slot] = index.flags[slot]! | DECLARED | (data.length > 12 && (f.u8(12) & 0x20) !== 0 ? HIDDEN : 0)
        }
        if (failure === undefined && data.length < ROW_MIN) fail(truncated(part, record), record.offset)
      } else if (type === STRING && failure === undefined && waitSlot !== -1) {
        // the formula's string result: BIFF8, an XLUnicodeString maybe
        // continued (the reader moved past its CONTINUEs), its character
        // count checked before any character; BIFF5, a u16-count code-page
        // string, never continued
        waited(record.offset, true)
        const read = globals.biff8
          ? readRichString(new RecordCursor(context.stream, record, context.end), limits.maxStringLength, false)
          : undefined
        const failed = read === undefined ? stringFailure(data, 0, globals.text, 2) : read._tag === "Ok" ? undefined : read
        const error = failed !== undefined ? stringFailed(failed, part, record, limits) : countCells(waitColumn, 1)
        if (error !== undefined) fail(error, record.offset)
      }
    }
    if (turn.spend(RECORD_COST)) yield
  }
  return { index, failure, failAt }
}

/**
 * The rows (see the module's comment): each row of the index in row order,
 * its records decoded into its cells, until a row wasn't complete before
 * the index pass's failure (then the failure). Rows go out through
 * `tally.rows`, in chunks of at most `BATCH_RECORDS` rows, one at the end of
 * each turn. A turn ends at a row's end once its budget is spent, or
 * within a row once it is `ROW_SLACK` past it.
 */
function* emitRows(context: SheetContext, indexed: Indexed, turn: Turn, tally: Tally): Generator<void, SpreadsheetError | undefined, void> {
  const { stream, globals, part, limits } = context
  const { index, failure, failAt } = indexed
  const { maxColumns, maxRows } = limits
  const order = index.ordered()
  const displaced = index.displaced
  const noted = index.noted
  /** The current row's next noted STRING (0 for none; see `formulaString`). */
  let note = 0
  const fields = new FixedReader()
  const row = new RowCells([])
  /** The rows taken so far are complete after this offset: the most of theirs. */
  let completeAt = -1
  /** Cells added to the current row, and cells sorted compacting it: the work since the last count. */
  let work = 0

  const add = (cell: Cell.Cell): void => {
    work += 1 + row.add(cell, maxColumns)
  }
  /** The text of the STRING record at `at`, as the index pass checked it. */
  const stringAt = (at: number): string | SpreadsheetError => {
    const record = { type: STRING, offset: at, data: stream.subarray(at + 4, at + 4 + (stream[at + 2]! | (stream[at + 3]! << 8))) }
    const read = globals.biff8
      ? readRichString(new RecordCursor(stream, record, context.end), limits.maxStringLength)
      : readString(record.data, 0, globals.text, 2)
    return read._tag === "Ok" ? read.value : stringFailed(read, part, record, limits)
  }
  /**
   * The string result of the FORMULA `record`: the STRING right after it,
   * "" if another record ends the wait there, or, when a SHRFMLA, ARRAY,
   * TABLE or CONTINUE comes first, the row's next STRING the index pass
   * noted (the row's FORMULAs are decoded in stream order, as noted).
   */
  const formulaString = (record: BiffRecord): string | SpreadsheetError => {
    const next = record.offset + 4 + record.data.length
    if (next + 4 > context.end) return ""
    const type = stream[next]! | (stream[next + 1]! << 8)
    if (type === STRING) return stringAt(next)
    if (!keepsWait(type) || note === 0) return ""
    const at = noted.values.items[note - 1]!
    note = noted.next.items[note - 1]!
    return at === 0 ? "" : stringAt(at)
  }
  /**
   * Decodes the cell record `record` of the current row into its cells.
   * The index pass checked it: a failure here is defensive.
   */
  const decode = (record: BiffRecord): SpreadsheetError | undefined => {
    const { data } = record
    switch (record.type) {
      case LABELSST: {
        const f = fields.fixed(data, LABELSST_MIN)
        if (f === undefined) return truncated(part, record)
        const value = globals.strings[f.u32(6)]
        if (value === undefined) return malformed(part, `shared string ${f.u32(6)} doesn't exist`, record.offset)
        add(Cell.text(f.u16(2), value))
        return undefined
      }
      case LABEL:
      case RSTRING: {
        const f = fields.fixed(data, CELL_MIN)
        if (f === undefined) return truncated(part, record)
        const column = f.u16(2)
        const text = readString(data, CELL_MIN, globals.text, 2)
        if (text._tag !== "Ok") return stringFailed(text, part, record, limits)
        add(Cell.text(column, text.value))
        return undefined
      }
      case NUMBER: {
        const f = fields.fixed(data, NUMBER_MIN)
        if (f === undefined) return truncated(part, record)
        add(xfNumber(globals, f.u16(2), f.u16(4), f.f64(6)))
        return undefined
      }
      case RK: {
        const f = fields.fixed(data, RK_MIN)
        if (f === undefined) return truncated(part, record)
        add(xfNumber(globals, f.u16(2), f.u16(4), decodeRk(f.u32(6))))
        return undefined
      }
      case MULRK:
      case MULBLANK: {
        const f = fields.fixed(data, MUL_MIN)
        if (f === undefined) return truncated(part, record)
        const first = f.u16(2)
        const mulrk = record.type === MULRK
        const width = mulrk ? 6 : 2
        const count = (data.length - MUL_MIN) / width
        // the cells fill the record between colFirst and colLast (checked by
        // the index pass): cell k's ixfe at 4 + width * k (and its rk after
        // it) ends by data.length - 2. progress: k up to count
        let sorted = 0
        for (let k = 0; k < count; k++) {
          const cell = mulrk ? xfNumber(globals, first + k, f.u16(4 + width * k), decodeRk(f.u32(6 + width * k))) : Cell.empty(first + k)
          sorted += row.add(cell, maxColumns)
        }
        work += count + sorted
        return undefined
      }
      case BOOLERR: {
        const f = fields.fixed(data, BOOLERR_MIN)
        if (f === undefined) return truncated(part, record)
        const column = f.u16(2)
        const value = f.u8(6)
        add(f.u8(7) === 1 ? Cell.error(column, errorCode(value)) : Cell.boolean(column, value !== 0))
        return undefined
      }
      case BLANK: {
        const f = fields.fixed(data, CELL_MIN)
        if (f === undefined) return truncated(part, record)
        add(Cell.empty(f.u16(2)))
        return undefined
      }
      default: {
        // FORMULA, FORMULA_BIFF4
        const f = fields.fixed(data, FORMULA_MIN)
        if (f === undefined) return truncated(part, record)
        const column = f.u16(2)
        if (f.u16(12) !== 0xFFFF) {
          add(xfNumber(globals, column, f.u16(4), f.f64(6)))
          return undefined
        }
        // a special result: string (in a STRING record), boolean, error, empty string
        const kind = f.u8(6)
        const value = f.u8(8)
        if (kind === 0) {
          const text = formulaString(record)
          if (typeof text !== "string") return text
          add(Cell.text(column, text))
        } else if (kind === 1) add(Cell.boolean(column, value !== 0))
        else if (kind === 2) add(Cell.error(column, errorCode(value)))
        else add(Cell.text(column, ""))
        return undefined
      }
    }
  }

  // progress: k up to order.length; within a row, the reader's offset grows
  // up to the range's end, and j up to the row's out-of-place records
  for (let k = 0; k < order.length; k++) {
    const slot = order[k]!
    const complete = index.complete[slot]!
    if (complete > completeAt) completeAt = complete
    // this row, or one before it, wasn't complete before the failure
    if (failure !== undefined && completeAt >= failAt) break
    const flags = index.flags[slot]!
    row.reset([])
    note = noted.head[slot]!
    if ((flags & RANGED) !== 0) {
      // the row's records that came together: cell records of other rows
      // don't come between them, other records may (embedded substreams too)
      const reader = new SubstreamReader(stream, part, index.start[slot]!, index.end[slot]!)
      for (let record = reader.read(); record !== undefined; record = reader.read()) {
        if (record instanceof MalformedFile) return record
        tally.steps++
        let cost = RECORD_COST
        if (!(record instanceof Embedded) && isCell(record.type)) {
          const error = decode(record)
          if (error !== undefined) return error
          tally.steps += work
          cost += work
          work = 0
        }
        if (turn.spend(cost) && turn.left <= -ROW_SLACK) yield
      }
    }
    // progress: each entry is later in the chain than the last (entries numbered in order), up to the slot's last
    for (let entry = displaced.head[slot]!; entry !== 0; entry = displaced.next.items[entry - 1]!) {
      const at = displaced.values.items[entry - 1]!
      const record = { type: stream[at]! | (stream[at + 1]! << 8), offset: at, data: stream.subarray(at + 4, at + 4 + (stream[at + 2]! | (stream[at + 3]! << 8))) }
      tally.steps++
      const error = decode(record)
      if (error !== undefined) return error
      tally.steps += work
      const cost = RECORD_COST + work
      work = 0
      if (turn.spend(cost) && turn.left <= -ROW_SLACK) yield
    }
    const sorting = row.sorted ? 0 : row.cells.length
    const cells = row.done()
    if (cells.length > 0 || (flags & DECLARED) !== 0) {
      const index_ = index.row[slot]!
      if (index_ >= maxRows) return new LimitExceeded({ limit: "maxRows", max: maxRows, actual: index_ + 1 })
      tally.rows.push({ index: index_, hidden: (flags & HIDDEN) !== 0, cells })
    }
    tally.steps += 1 + sorting
    if (turn.spend(1 + sorting) || tally.rows.length >= BATCH_RECORDS) yield
  }
  return failure
}

/** The rows walk: open the substream, the index pass, then the rows. */
function* rowsWalk(context: SheetContext, turn: Turn, tally: Tally): Generator<void, SpreadsheetError | undefined, void> {
  // a BIFF8 STRING comes with its CONTINUEs (a BIFF5 one has none)
  const opened = openSubstream(context, context.globals.biff8 ? STRING : undefined)
  if (!(opened instanceof SubstreamReader)) return opened
  const indexed = yield* indexRecords(context, opened, turn, tally)
  return yield* emitRows(context, indexed, turn, tally)
}

/**
 * The rows of a worksheet substream (see the module's comment), in turns of
 * at most `BATCH_RECORDS` records or about `BATCH_CELLS` cells, with a
 * yield to the event loop before each. Re-runnable.
 */
export const sheetRows = (context: SheetContext): Stream.Stream<Row, SpreadsheetError> =>
  Stream.fromPull(Effect.map(WorkMeter, (meter) => {
    const turn = new Turn(BATCH_CELLS)
    const tally: Tally = { steps: 0, rows: [] }
    const walk = rowsWalk(context, turn, tally)
    /** Hands the steps of the last turn to the meter. */
    const report = Effect.sync(() => {
      meter.add("steps", tally.steps)
      tally.steps = 0
    })
    /** The walk's end: running (null), done (undefined) or its failure. */
    let ended: SpreadsheetError | undefined | null = null
    const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, SpreadsheetError> = Effect.gen(function*() {
      // progress: each iteration runs the walk for one turn (each turn
      // advances it), or hands out rows, or ends the stream
      while (true) {
        yield* report
        yield* yieldNow
        if (ended === null) {
          turn.left = turn.size
          const step = walk.next()
          if (step.done === true) ended = step.value
        }
        // the rows of the turn go out first, also those before a failure
        if (tally.rows.length > 0) {
          const rows = tally.rows as unknown as Arr.NonEmptyReadonlyArray<Row>
          tally.rows = []
          return rows
        }
        if (ended === undefined) return yield* Cause.done()
        if (ended !== null) return yield* Effect.fail(ended)
      }
    })
    // the steps of the last turn count also when the pull fails
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

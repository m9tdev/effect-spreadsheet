/**
 * A worksheet part of an XLSB package -> rows, streamed; and its layout
 * (`sheetDetails`).
 *
 * The part is inflated and split into records piece by piece (an inflated
 * piece is at most 64 KiB, so at most 32 768 records), and the rows a
 * piece completes are emitted with it; the stream yields to the event
 * loop before each piece. So the work and the rows between event-loop
 * turns are bounded.
 *
 * - Rows: each BrtRowHdr starts a row (index, hidden flag `fDyZero`); its
 *   cell records follow it. A row is emitted when the next row starts or
 *   the sheet data ends, in document order; rows without cells are emitted
 *   too, rows the file leaves out are not. Hidden rows are emitted with
 *   `hidden: true`, never skipped.
 * - Cells: blank, RK, error, boolean, real, inline string, rich string and
 *   shared string cells, the four formula cells (their cached result; the
 *   formula isn't decompiled), and the short forms Excel 2010 and later
 *   may write (no column field: the previous cell's column + 1). NaN and
 *   infinities are `#NUM!`, unknown error codes `#N/A`, as in XLS. A number
 *   whose style has a date format is a `Date`, unless it is outside
 *   Excel's date range; numbers carry their `numFmt`.
 * - Cells are held only for the open row. A row whose cells are out of
 *   order is sorted, keeping the last of duplicates, and compacted that way
 *   as soon as it holds more than 2 x `maxColumns` cells (as in XLSX), so
 *   the cells held stay under that whatever the file repeats.
 * - Nothing is sized by BrtWsDim; cost follows the records present.
 *
 * @internal
 */
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import * as Cell from "../../Cell.ts"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { ColumnSpan, Range, Row, SheetDetails } from "../../Workbook.ts"
import { Fields, readRichStr, readWideString } from "../biff12/fields.ts"
import { type Biff12Record, readRecords, RecordSplitter } from "../biff12/records.ts"
import { decodeRk, ERROR_CODES } from "../xls/fields.ts"
import { mergeSpans } from "../xlsx/details.ts"
import { serialToDate } from "../xlsx/numFmt.ts"
import type { Styles } from "../xlsx/parts.ts"
import { compact } from "../xlsx/sheet.ts"
import { malformed } from "./workbook.ts"

export interface SheetContext {
  readonly part: string
  readonly strings: ReadonlyArray<string> | undefined
  readonly styles: Styles
  readonly date1904: boolean
  readonly limits: Limits
}

/**
 * Records handled by the rows reader and by sheetDetails, summed over every
 * read in this process. Tests use it to check that the work grows linearly
 * with the sheet, whatever the machine's load.
 */
export const work = { records: 0 }

const BRT_ROW_HDR = 0x0000
const BRT_CELL_BLANK = 0x0001
const BRT_CELL_RK = 0x0002
const BRT_CELL_ERROR = 0x0003
const BRT_CELL_BOOL = 0x0004
const BRT_CELL_REAL = 0x0005
const BRT_CELL_ST = 0x0006
const BRT_CELL_ISST = 0x0007
const BRT_FMLA_STRING = 0x0008
const BRT_FMLA_NUM = 0x0009
const BRT_FMLA_BOOL = 0x000A
const BRT_FMLA_ERROR = 0x000B
const BRT_SHORT_BLANK = 0x000C
const BRT_SHORT_RK = 0x000D
const BRT_SHORT_ERROR = 0x000E
const BRT_SHORT_BOOL = 0x000F
const BRT_SHORT_REAL = 0x0010
const BRT_SHORT_ST = 0x0011
const BRT_SHORT_ISST = 0x0012
const BRT_CELL_RSTRING = 0x003E
const BRT_SHORT_RSTRING = 0x003F
const BRT_END_SHEET_DATA = 0x0092
const BRT_WS_DIM = 0x0094
const BRT_COL_INFO = 0x003C
const BRT_MERGE_CELL = 0x00B0

/** The row header's fDyZero bit (in the u16 at offset 10): a hidden row. */
const ROW_HIDDEN = 0x1000

const BRT_BEGIN_SHEET_DATA = 0x0091
const BRT_FRT_BEGIN = 0x0023
const BRT_FRT_END = 0x0024

/**
 * Where a record sits: record numbers are reused inside future-record
 * blocks (BrtFRTBegin ... BrtFRTEnd, e.g. a record 0x0001 that isn't a
 * blank cell), so only records outside them count; and row and cell
 * records count only inside the sheet data. The depth is bounded by the
 * records read (each BrtFRTBegin is a record).
 */
class Scope {
  /** Open future-record blocks. */
  private frt = 0
  /** Inside BrtBeginSheetData ... BrtEndSheetData. */
  sheetData = false

  /** Takes `type` into account; true if it is a record of the sheet itself (outside future-record blocks). */
  step(type: number): boolean {
    if (type === BRT_FRT_BEGIN) {
      this.frt++
      return false
    }
    if (this.frt > 0) {
      if (type === BRT_FRT_END) this.frt--
      return false
    }
    if (type === BRT_BEGIN_SHEET_DATA) this.sheetData = true
    else if (type === BRT_END_SHEET_DATA) {
      const was = this.sheetData
      this.sheetData = false
      return was
    }
    return this.sheetData || (type !== BRT_ROW_HDR && !(type <= BRT_SHORT_ISST || type === BRT_CELL_RSTRING || type === BRT_SHORT_RSTRING))
  }
}

const errorCode = (code: number): string => ERROR_CODES.get(code) ?? "#N/A"

/** Collects rows from a worksheet's records. Exported for tests. */
export class SheetRows {
  /** Completed rows not yet emitted. */
  out: Array<Row> = []
  private cells = 0
  private rowCount = 0
  /** The open row's index; undefined before the first row and after the sheet data. */
  private rowIndex: number | undefined
  private rowHidden = false
  private rowCells: Array<Cell.Cell> = []
  private rowSorted = true
  /** The open row's last cell column (short cells follow it), -1 before its first. */
  private lastColumn = -1
  private readonly context: SheetContext

  constructor(context: SheetContext) {
    this.context = context
  }

  /** Cells held for the open row (for tests). */
  get heldCells(): number {
    return this.rowCells.length
  }

  private readonly scope = new Scope()

  /** Handles one record. */
  handle = (record: Biff12Record): SpreadsheetError | undefined => {
    work.records++
    const type = record.type
    if (!this.scope.step(type)) return undefined
    if (type === BRT_ROW_HDR) return this.openRow(record)
    if (type === BRT_END_SHEET_DATA) return this.closeRow()
    if (type <= BRT_SHORT_ISST || type === BRT_CELL_RSTRING || type === BRT_SHORT_RSTRING) return this.cell(record)
    return undefined
  }

  /** The end of the part: the open row, if any, is complete. */
  end = (): SpreadsheetError | undefined => this.closeRow()

  private openRow(record: Biff12Record): SpreadsheetError | undefined {
    const error = this.closeRow()
    if (error !== undefined) return error
    const limits = this.context.limits
    // rw u32, ixfe u32, miyRw u16, flags u16 (fDyZero), ...
    const f = new Fields(record.data)
    const index = f.u32(0)
    const flags = f.u16(10)
    if (index === undefined || flags === undefined) return malformed("truncated row header", this.context.part, record.offset)
    if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
    if (++this.rowCount > limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: this.rowCount })
    this.rowIndex = index
    this.rowHidden = (flags & ROW_HIDDEN) !== 0
    this.rowCells = []
    this.rowSorted = true
    this.lastColumn = -1
    return undefined
  }

  private closeRow(): undefined {
    if (this.rowIndex === undefined) return undefined
    const cells = this.rowSorted ? this.rowCells : compact(this.rowCells)
    this.out.push({ index: this.rowIndex, hidden: this.rowHidden, cells })
    this.rowIndex = undefined
    this.rowCells = []
    return undefined
  }

  private cell(record: Biff12Record): SpreadsheetError | undefined {
    const { part, limits } = this.context
    if (this.rowIndex === undefined) return malformed(`cell record 0x${record.type.toString(16)} outside a row`, part, record.offset)
    const type = record.type
    const data = record.data
    const f = new Fields(data)
    const short = (type >= BRT_SHORT_BLANK && type <= BRT_SHORT_ISST) || type === BRT_SHORT_RSTRING
    // Cell: column u32 and style u24 (+ flags byte); the short forms have no column
    const column = short ? this.lastColumn + 1 : f.u32(0)
    const style = f.u32(short ? 0 : 4)
    const at = short ? 4 : 8
    const truncated = () => malformed(`truncated cell record 0x${type.toString(16)}`, part, record.offset)
    if (column === undefined || style === undefined) return truncated()
    if (column >= limits.maxColumns) return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: column + 1 })
    let cell: Cell.Cell
    switch (type) {
      case BRT_CELL_BLANK:
      case BRT_SHORT_BLANK:
        cell = Cell.empty(column)
        break
      case BRT_CELL_RK:
      case BRT_SHORT_RK: {
        const rk = f.u32(at)
        if (rk === undefined) return truncated()
        cell = this.numberCell(column, style & 0xFFFFFF, decodeRk(rk))
        break
      }
      case BRT_CELL_REAL:
      case BRT_SHORT_REAL:
      case BRT_FMLA_NUM: {
        const value = f.f64(at)
        if (value === undefined) return truncated()
        cell = this.numberCell(column, style & 0xFFFFFF, value)
        break
      }
      case BRT_CELL_BOOL:
      case BRT_SHORT_BOOL:
      case BRT_FMLA_BOOL: {
        const value = f.u8(at)
        if (value === undefined) return truncated()
        cell = Cell.boolean(column, value !== 0)
        break
      }
      case BRT_CELL_ERROR:
      case BRT_SHORT_ERROR:
      case BRT_FMLA_ERROR: {
        const value = f.u8(at)
        if (value === undefined) return truncated()
        cell = Cell.error(column, errorCode(value))
        break
      }
      case BRT_CELL_ST:
      case BRT_SHORT_ST:
      case BRT_FMLA_STRING:
      case BRT_CELL_RSTRING:
      case BRT_SHORT_RSTRING: {
        const rich = type === BRT_CELL_RSTRING || type === BRT_SHORT_RSTRING
        const text = rich ? readRichStr(data, at, limits.maxStringLength) : readWideString(data, at, limits.maxStringLength)
        if (text === undefined) return truncated()
        if (text === null) return new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual: limits.maxStringLength + 1 })
        cell = Cell.text(column, text.value)
        break
      }
      case BRT_CELL_ISST:
      case BRT_SHORT_ISST: {
        // a shared string index
        const index = f.u32(at)
        if (index === undefined) return truncated()
        const strings = this.context.strings
        if (strings === undefined || index >= strings.length) return malformed(`shared string ${index} doesn't exist`, part, record.offset)
        cell = Cell.text(column, strings[index]!)
        break
      }
      default:
        return undefined
    }
    if (++this.cells > limits.maxCellsPerSheet) {
      return new LimitExceeded({ limit: "maxCellsPerSheet", max: limits.maxCellsPerSheet, actual: this.cells })
    }
    this.lastColumn = column
    const cells = this.rowCells
    if (cells.length > 0 && column <= cells[cells.length - 1]!.column) this.rowSorted = false
    cells.push(cell)
    // repeated columns don't pile up: past 2 x maxColumns cells an unsorted
    // row is compacted (to at most maxColumns), so the work is amortized
    if (!this.rowSorted && cells.length > 2 * limits.maxColumns) {
      compact(cells)
      this.rowSorted = true
    }
    return undefined
  }

  /** A number cell, a Date when its style has a date format. */
  private numberCell(column: number, style: number, value: number): Cell.Cell {
    const format = this.context.styles.formats[style]
    const meta: Cell.Meta | undefined = format?.code === undefined ? undefined : { numFmt: format.code }
    if (format?.date === true) {
      const date = serialToDate(value, this.context.date1904)
      if (date !== undefined) return Cell.date(column, date, meta)
    }
    // NaN and infinities aren't Excel numbers: #NUM!, as for a formula result
    if (!Number.isFinite(value)) return Cell.error(column, "#NUM!", meta)
    return Cell.number(column, value === 0 ? 0 : value, meta)
  }
}

/** Rows of a worksheet part, given its byte stream. Re-runnable when the input is. */
export const sheetRows = (
  bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
  context: SheetContext
): Stream.Stream<Row, SpreadsheetError> =>
  Stream.transformPull(bytes, (upstream) =>
    Effect.sync(() => {
      const rows = new SheetRows(context)
      const splitter = new RecordSplitter(context.part, context.limits.maxRecordBytes)
      const pullUpstream = Pull.catchDone(upstream, () => Effect.succeed(undefined))
      let done = false
      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, SpreadsheetError> = Effect.gen(function*() {
        // progress: each iteration pulls one upstream piece (consumed), or
        // ends the input (done is set), or ends the stream
        while (true) {
          if (done) return yield* Cause.done()
          yield* Effect.yieldNow
          const chunk = yield* pullUpstream
          let error: SpreadsheetError | undefined
          if (chunk === undefined) {
            done = true
            error = splitter.end() ?? rows.end()
          } else {
            // progress: one piece per iteration
            for (const piece of chunk) {
              error = splitter.feed(piece, rows.handle)
              if (error !== undefined) break
            }
          }
          if (error !== undefined) {
            done = true
            return yield* Effect.fail(error)
          }
          if (rows.out.length > 0) {
            const out = rows.out
            rows.out = []
            return out as unknown as Arr.NonEmptyReadonlyArray<Row>
          }
        }
      })
      return pull
    }))

/** An RfX range (rwFirst, rwLast, colFirst, colLast: u32 each) at `at`; undefined if truncated or reversed. */
const range = (f: Fields, at: number): Range | undefined => {
  const firstRow = f.u32(at)
  const lastRow = f.u32(at + 4)
  const firstColumn = f.u32(at + 8)
  const lastColumn = f.u32(at + 12)
  if (firstRow === undefined || lastRow === undefined || firstColumn === undefined || lastColumn === undefined) return undefined
  if (lastRow < firstRow || lastColumn < firstColumn) return undefined
  return { firstRow, firstColumn, lastRow, lastColumn }
}

/**
 * Hidden columns (BrtColInfo), merged ranges (BrtMergeCell) and the
 * declared dimension (BrtWsDim) of a worksheet part, in one pass over its
 * records; cell records are split off but not decoded. Merges come after
 * the sheet data, so the whole part is read.
 */
export const sheetDetails = (
  bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
  part: string,
  limits: Limits
): Effect.Effect<SheetDetails, SpreadsheetError> =>
  Effect.suspend(() => {
    const hidden: Array<ColumnSpan> = []
    const merges: Array<Range> = []
    let colInfos = 0
    let dimension: Range | undefined
    const scope = new Scope()
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      work.records++
      if (!scope.step(record.type)) return undefined
      switch (record.type) {
        case BRT_WS_DIM: {
          const r = range(new Fields(record.data), 0)
          if (r !== undefined) dimension = r
          return undefined
        }
        case BRT_COL_INFO: {
          if (++colInfos > limits.maxColumns) return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: colInfos })
          // colFirst u32, colLast u32, coldx u32, ixfe u32, flags u16 (fHidden bit 0)
          const f = new Fields(record.data)
          const first = f.u32(0)
          const last = f.u32(4)
          const flags = f.u16(16)
          if (first === undefined || last === undefined || flags === undefined) return malformed("truncated column info", part, record.offset)
          if ((flags & 0x01) === 0) return undefined
          const end = Math.min(last, limits.maxColumns - 1)
          if (first <= end) hidden.push({ first, last: end })
          return undefined
        }
        case BRT_MERGE_CELL: {
          const r = range(new Fields(record.data), 0)
          if (r === undefined) return malformed("bad merge range", part, record.offset)
          if (merges.length >= limits.maxCellsPerSheet) {
            return new LimitExceeded({ limit: "maxCellsPerSheet", max: limits.maxCellsPerSheet, actual: merges.length + 1 })
          }
          merges.push(r)
          return undefined
        }
        default:
          return undefined
      }
    }
    return Effect.map(readRecords(bytes, handle, limits, part), (): SheetDetails => ({
      hiddenColumns: mergeSpans(hidden),
      merges,
      ...(dimension === undefined ? {} : { dimension })
    }))
  })

/**
 * A worksheet part of an XLSB package -> rows, streamed; and its layout
 * (`sheetDetails`).
 *
 * The part is inflated and split into records slice by slice, and the
 * rows a slice completes are emitted with it. The archive hands the part
 * over in slices of at most `SLICE_BYTES` (64 KiB, so at most 32 768
 * records), each after a yield to the event loop, so the work and the
 * rows between event-loop turns are bounded, however the decompressor
 * cuts its output.
 *
 * - Rows: each BrtRowHdr starts a row (index, hidden flag `fDyZero`); its
 *   cell records follow it. A row is emitted when the next row starts or
 *   the sheet data ends, in document order; rows without cells are emitted
 *   too, rows the file leaves out are not. Hidden rows are emitted with
 *   `hidden: true`, never skipped. Row indexes must go up: a BrtRowHdr whose
 *   index is the same as or below the previous row's is MalformedFile (rows
 *   are streamed, so a row already emitted can't be merged with a later one).
 * - Cells: blank, RK, error, boolean, real, inline string, rich string and
 *   shared string cells, the four formula cells (their cached result; the
 *   formula isn't decompiled), and the short forms MS-XLSB defines (no
 *   column field: the previous cell's column + 1 in the row, 0 for the
 *   row's first cell). NaN and
 *   infinities are `#NUM!`, unknown error codes `#N/A`, as in XLS. A number
 *   whose style has a date format is a `Date`, unless it is outside
 *   Excel's date range; numbers carry their `numFmt`.
 * - Cells are held only for the open row. A row whose cells are out of
 *   order is sorted, keeping the last of duplicates, and compacted that way
 *   as soon as it holds more than 2 x `maxColumns` cells (as in XLSX), so
 *   the cells held stay under that whatever the file repeats.
 * - Nothing is sized by BrtWsDim; cost follows the records present.
 * - On an error (malformed input, a limit), the rows completed before it
 *   come out first, then the stream fails with it (as in XLSX). So the rows
 *   seen before the error don't depend on where the slices fall.
 *
 * @internal
 */
import { Effect, type Stream } from "effect"
import * as Cell from "../../Cell.ts"
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Range, Row, SheetDetails } from "../../Workbook.ts"
import { decodeRk, ERROR_CODES, Fields } from "../binary.ts"
import { readRichStr, readWideString } from "../biff12/fields.ts"
import { type Biff12Record, readRecords, recordDecoder } from "../biff12/records.ts"
import { decoded, step } from "../decode.ts"
import { type CellStyles, numberCell } from "../numFmt.ts"
import { LayoutBuilder } from "../sheet/layout.ts"
import { IGNORED, RowBuilder } from "../sheet/rows.ts"
import { malformed, tooLong } from "./workbook.ts"

export interface SheetContext {
  readonly part: string
  readonly strings: ReadonlyArray<string> | undefined
  readonly styles: CellStyles
  readonly date1904: boolean
  readonly limits: Limits
}

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
 * records read (each BrtFRTBegin is a record). A block still open at the
 * end of the part would hide every record after its BrtFRTBegin, so it is
 * malformed (`end`).
 */
class Scope {
  /** Open future-record blocks. */
  private frt = 0
  /** Offset of the outermost open block's BrtFRTBegin. */
  private frtOffset = 0
  /** Inside BrtBeginSheetData ... BrtEndSheetData. */
  sheetData = false

  /** Takes `record` into account; true if it is a record of the sheet itself (outside future-record blocks). */
  step(record: Biff12Record): boolean {
    const type = record.type
    if (type === BRT_FRT_BEGIN) {
      if (this.frt++ === 0) this.frtOffset = record.offset
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
    return this.sheetData || (type !== BRT_ROW_HDR && !(type <= BRT_SHORT_ISST || type === BRT_CELL_RSTRING))
  }

  /** The end of the part: MalformedFile if a future-record block is still open. */
  end(part: string): SpreadsheetError | undefined {
    return this.frt > 0 ? malformed("future-record block not closed", part, this.frtOffset) : undefined
  }
}

const errorCode = (code: number): string => ERROR_CODES.get(code) ?? "#N/A"

/** Collects rows from a worksheet's records. Exported for tests. */
export class SheetRows {
  /** The rows: in XLSB a cell record outside a row is MalformedFile. */
  readonly rows: RowBuilder<Biff12Record>
  private readonly context: SheetContext

  constructor(context: SheetContext) {
    this.context = context
    const part = context.part
    this.rows = new RowBuilder<Biff12Record>(context.limits, {
      cellOutsideRow: (record) => malformed(`cell record 0x${record.type.toString(16)} outside a row`, part, record.offset),
      malformed: (reason, record) => malformed(reason, part, record.offset)
    })
  }

  /** The completed rows not yet taken. */
  take(): ReadonlyArray<Row> {
    return this.rows.take()
  }

  private readonly scope = new Scope()

  /** Handles one record. */
  handle = (record: Biff12Record): SpreadsheetError | undefined => {
    const type = record.type
    if (!this.scope.step(record)) return undefined
    if (type === BRT_ROW_HDR) return this.openRow(record)
    if (type === BRT_END_SHEET_DATA) return this.rows.closeRow()
    if (type <= BRT_SHORT_ISST || type === BRT_CELL_RSTRING) return this.cell(record)
    return undefined
  }

  /** The end of the part: the open row, if any, is complete (MalformedFile if a future-record block is still open). */
  end = (): SpreadsheetError | undefined => this.scope.end(this.context.part) ?? this.rows.closeRow()

  /** A BrtRowHdr: the open row is complete (even if this header is truncated), and the next one opens. */
  private openRow(record: Biff12Record): SpreadsheetError | undefined {
    const closed = this.rows.closeRow()
    if (closed !== undefined) return closed
    // rw u32, ixfe u32, miyRw u16, flags u16 (fDyZero), ...
    const f = new Fields(record.data)
    const index = f.u32(0)
    const flags = f.u16(10)
    if (index === undefined || flags === undefined) return malformed("truncated row header", this.context.part, record.offset)
    return this.rows.openRow(index, (flags & ROW_HIDDEN) !== 0, record)
  }

  private cell(record: Biff12Record): SpreadsheetError | undefined {
    const { part, limits } = this.context
    const rows = this.rows
    const outside = rows.beginCell(record)
    if (outside !== undefined) return outside === IGNORED ? undefined : outside
    const type = record.type
    const data = record.data
    const f = new Fields(data)
    const short = type >= BRT_SHORT_BLANK && type <= BRT_SHORT_ISST
    // Cell: column u32 and style u24 (+ flags byte); the short forms have no column
    const column = short ? rows.nextColumn : f.u32(0)
    const style = f.u32(short ? 0 : 4)
    const at = short ? 4 : 8
    const truncated = () => malformed(`truncated cell record 0x${type.toString(16)}`, part, record.offset)
    if (column === undefined || style === undefined) return truncated()
    const placed = rows.placeCell(column)
    if (placed !== undefined) return placed
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
      case BRT_CELL_RSTRING: {
        const rich = type === BRT_CELL_RSTRING
        const text = rich ? readRichStr(data, at, limits.maxStringLength) : readWideString(data, at, limits.maxStringLength)
        if (text === undefined) return truncated()
        if (typeof text === "number") return tooLong(limits, text)
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
    return rows.addCell(cell)
  }

  /** A number cell, a Date when its style has a date format. */
  private numberCell(column: number, style: number, value: number): Cell.Cell {
    return numberCell(column, value, this.context.styles.format(style), this.context.date1904)
  }
}

/**
 * Rows of a worksheet part, given its byte stream. Re-runnable when the
 * input is. The records of each slice go to the work meter.
 */
export const sheetRows = (
  bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
  context: SheetContext
): Stream.Stream<Row, SpreadsheetError> =>
  decoded(bytes, (meter) => {
    const rows = new SheetRows(context)
    const records = recordDecoder(rows.handle, context.limits, context.part, meter)
    return {
      feed: (piece) => {
        const error = records.feed(piece).error
        return step(rows.take(), error)
      },
      end: () => {
        const error = records.end().error ?? rows.end()
        return step(rows.take(), error)
      }
    }
  })

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
    const layout = new LayoutBuilder(limits)
    const scope = new Scope()
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      if (!scope.step(record)) return undefined
      switch (record.type) {
        case BRT_WS_DIM: {
          const r = range(new Fields(record.data), 0)
          if (r !== undefined) layout.dimension(r)
          return undefined
        }
        case BRT_COL_INFO: {
          const counted = layout.columnRecord()
          if (counted !== undefined) return counted
          // colFirst u32, colLast u32, coldx u32, ixfe u32, flags u16 (fHidden bit 0)
          const f = new Fields(record.data)
          const first = f.u32(0)
          const last = f.u32(4)
          const flags = f.u16(16)
          if (first === undefined || last === undefined || flags === undefined) return malformed("truncated column info", part, record.offset)
          if ((flags & 0x01) !== 0) layout.hideColumns(first, last)
          return undefined
        }
        case BRT_MERGE_CELL: {
          const r = range(new Fields(record.data), 0)
          if (r === undefined) return malformed("bad merge range", part, record.offset)
          return layout.merge(r)
        }
        default:
          return undefined
      }
    }
    return Effect.flatMap(readRecords(bytes, handle, limits, part), () => {
      const error = scope.end(part)
      if (error !== undefined) return Effect.fail(error)
      return Effect.succeed(layout.result())
    })
  })

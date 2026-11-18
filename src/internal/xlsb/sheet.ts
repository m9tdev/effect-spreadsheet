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
import { decodeRk, ERROR_CODES } from "../binary.ts"
import { type Fixed, FixedReader, type StringFailure, type StringRead } from "../binary/layout.ts"
import { readRichStr, readWideString } from "../biff12/fields.ts"
import { type Biff12Record, readRecords, recordDecoder } from "../biff12/records.ts"
import { decoded, step } from "../decode.ts"
import { type CellStyles, numberCell } from "../numFmt.ts"
import { LayoutBuilder } from "../sheet/layout.ts"
import { RowBuilder } from "../sheet/rows.ts"
import { malformed, stringError } from "./workbook.ts"

export interface SheetContext {
  readonly part: string
  readonly strings: ReadonlyArray<string> | undefined
  readonly styles: CellStyles
  readonly date1904: boolean
  readonly limits: Limits
}

const BRT_ROW_HDR = 0x0000
const BRT_END_SHEET_DATA = 0x0092
const BRT_WS_DIM = 0x0094
const BRT_COL_INFO = 0x003C
const BRT_MERGE_CELL = 0x00B0

/** The row header's fDyZero bit (in the u16 at offset 10): a hidden row. */
const ROW_HIDDEN = 0x1000
/** A row header's fixed fields: rw u32, ixfe u32, miyRw u16, flags u16 (fDyZero), ... */
const ROW_HDR_MIN = 12
/** An RfX range: rwFirst, rwLast, colFirst, colLast, u32 each. */
const RFX_MIN = 16
/** A BrtColInfo's fixed fields: colFirst u32, colLast u32, coldx u32, ixfe u32, flags u16 (fHidden bit 0). */
const COL_INFO_MIN = 18

const errorCode = (code: number): string => ERROR_CODES.get(code) ?? "#N/A"

/** Where a cell record's value goes: the sheet's rows, read in its context. */
interface CellSink {
  readonly context: SheetContext
  /** Adds the cell to the open row. */
  add(cell: Cell.Cell): SpreadsheetError | undefined
  /** Adds a number cell, a Date when its style has a date format. */
  addNumber(column: number, style: number, value: number): SpreadsheetError | undefined
  /** The error of a cell's string that couldn't be read: too long, or past the record. */
  unread(read: StringFailure, record: Biff12Record): SpreadsheetError
}

/** Decodes a cell's value at `at` (after its column and style) and adds the cell. */
type CellValue = (sink: CellSink, f: Fixed, at: number, column: number, style: number, record: Biff12Record) => SpreadsheetError | undefined

/**
 * A cell record type: whether it is a short form (no column field: the
 * previous cell's column + 1), its declared minimum length (the column, if
 * any, and style u24 + flags byte, then the value's fixed part; a string
 * value checks its own length), and how its value is read.
 */
interface CellKind {
  readonly short: boolean
  readonly min: number
  readonly value: CellValue
}

const blankValue: CellValue = (sink, _f, _at, column) => sink.add(Cell.empty(column))
const rkValue: CellValue = (sink, f, at, column, style) => sink.addNumber(column, style, decodeRk(f.u32(at)))
const realValue: CellValue = (sink, f, at, column, style) => sink.addNumber(column, style, f.f64(at))
const boolValue: CellValue = (sink, f, at, column) => sink.add(Cell.boolean(column, f.u8(at) !== 0))
const errorValue: CellValue = (sink, f, at, column) => sink.add(Cell.error(column, errorCode(f.u8(at))))
const textOf = (read: StringRead, sink: CellSink, column: number, record: Biff12Record): SpreadsheetError | undefined =>
  read._tag === "Ok" ? sink.add(Cell.text(column, read.value)) : sink.unread(read, record)
const textValue: CellValue = (sink, f, at, column, _style, record) =>
  textOf(readWideString(f.data, at, sink.context.limits.maxStringLength), sink, column, record)
const richValue: CellValue = (sink, f, at, column, _style, record) =>
  textOf(readRichStr(f.data, at, sink.context.limits.maxStringLength), sink, column, record)
const sharedValue: CellValue = (sink, f, at, column, _style, record) => {
  // a shared string index
  const index = f.u32(at)
  const strings = sink.context.strings
  if (strings === undefined || index >= strings.length) return malformed(`shared string ${index} doesn't exist`, sink.context.part, record.offset)
  return sink.add(Cell.text(column, strings[index]!))
}

const truncatedCell = (record: Biff12Record): string => `truncated cell record 0x${record.type.toString(16)}`

const cellKind = (short: boolean, size: number, value: CellValue): CellKind => ({ short, min: (short ? 4 : 8) + size, value })

/**
 * The cell records, by record type: they drive which records are cells
 * (inside the sheet data) and how each is read. Types 0x01 to 0x12 and
 * 0x3E; none above.
 */
const CELLS: ReadonlyArray<CellKind | undefined> = (() => {
  const table: Array<CellKind | undefined> = new Array(0x3F).fill(undefined)
  table[0x01] = cellKind(false, 0, blankValue) // BrtCellBlank
  table[0x02] = cellKind(false, 4, rkValue) // BrtCellRk
  table[0x03] = cellKind(false, 1, errorValue) // BrtCellError
  table[0x04] = cellKind(false, 1, boolValue) // BrtCellBool
  table[0x05] = cellKind(false, 8, realValue) // BrtCellReal
  table[0x06] = cellKind(false, 0, textValue) // BrtCellSt
  table[0x07] = cellKind(false, 4, sharedValue) // BrtCellIsst
  table[0x08] = cellKind(false, 0, textValue) // BrtFmlaString
  table[0x09] = cellKind(false, 8, realValue) // BrtFmlaNum
  table[0x0A] = cellKind(false, 1, boolValue) // BrtFmlaBool
  table[0x0B] = cellKind(false, 1, errorValue) // BrtFmlaError
  table[0x0C] = cellKind(true, 0, blankValue) // BrtShortBlank
  table[0x0D] = cellKind(true, 4, rkValue) // BrtShortRk
  table[0x0E] = cellKind(true, 1, errorValue) // BrtShortError
  table[0x0F] = cellKind(true, 1, boolValue) // BrtShortBool
  table[0x10] = cellKind(true, 8, realValue) // BrtShortReal
  table[0x11] = cellKind(true, 0, textValue) // BrtShortSt
  table[0x12] = cellKind(true, 4, sharedValue) // BrtShortIsst
  table[0x3E] = cellKind(false, 0, richValue) // BrtCellRString
  return table
})()

/** The kind of a cell record type, if it is one. */
const cellKindOf = (type: number): CellKind | undefined => (type < CELLS.length ? CELLS[type] : undefined)

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
    return this.sheetData || (type !== BRT_ROW_HDR && cellKindOf(type) === undefined)
  }

  /** The end of the part: MalformedFile if a future-record block is still open. */
  end(part: string): SpreadsheetError | undefined {
    return this.frt > 0 ? malformed("future-record block not closed", part, this.frtOffset) : undefined
  }
}

/** Collects rows from a worksheet's records. Exported for tests. */
export class SheetRows implements CellSink {
  /** The rows: in XLSB a cell record outside a row is MalformedFile. */
  readonly rows: RowBuilder<Biff12Record>
  readonly context: SheetContext
  private readonly reader = new FixedReader()

  constructor(context: SheetContext) {
    this.context = context
    const part = context.part
    this.rows = new RowBuilder<Biff12Record>(context.limits, {
      cellOutsideRow: (record) => malformed(`cell record 0x${record.type.toString(16)} outside a row`, part, record.offset),
      malformed: (reason, record) => malformed(reason, part, record.offset),
      // cell records don't name their row: never reached
      cellInOtherRow: (row, record) => malformed(`cell record 0x${record.type.toString(16)} outside row ${row + 1}`, part, record.offset)
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
    const kind = cellKindOf(type)
    return kind === undefined ? undefined : this.cell(record, kind)
  }

  /** The end of the part: the open row, if any, is complete (MalformedFile if a future-record block is still open). */
  end = (): SpreadsheetError | undefined => this.scope.end(this.context.part) ?? this.rows.closeRow()

  /** A BrtRowHdr: the open row is complete (even if this header is truncated), and the next one opens. */
  private openRow(record: Biff12Record): SpreadsheetError | undefined {
    const closed = this.rows.closeRow()
    if (closed !== undefined) return closed
    const f = this.reader.fixed(record.data, ROW_HDR_MIN)
    if (f === undefined) return malformed("truncated row header", this.context.part, record.offset)
    return this.rows.openRow(f.u32(0), (f.u16(10) & ROW_HIDDEN) !== 0, record)
  }

  /**
   * A cell record of `kind`. Its column and style are checked and read
   * first, then the cell is placed (which may fail on its own), and only
   * then is the record checked against the kind's declared minimum and its
   * value read.
   */
  private cell(record: Biff12Record, kind: CellKind): SpreadsheetError | undefined {
    const rows = this.rows
    const start = rows.beginCell(record)
    if (start !== "read") return start === "skip" ? undefined : start
    const data = record.data
    // Cell: column u32 and style u24 (+ flags byte); the short forms have no column
    const head = kind.short ? 4 : 8
    const f = this.reader.fixed(data, head)
    if (f === undefined) return this.truncated(record)
    const column = kind.short ? rows.nextColumn : f.u32(0)
    const style = f.u32(head - 4) & 0xFFFFFF
    const placed = rows.placeCell(column, record)
    if (placed !== undefined) return placed
    if (data.length < kind.min) return this.truncated(record)
    return kind.value(this, f, head, column, style, record)
  }

  private truncated(record: Biff12Record): SpreadsheetError {
    return malformed(truncatedCell(record), this.context.part, record.offset)
  }

  unread(read: StringFailure, record: Biff12Record): SpreadsheetError {
    return stringError(read, this.context.limits, truncatedCell(record), this.context.part, record.offset)
  }

  add(cell: Cell.Cell): SpreadsheetError | undefined {
    return this.rows.addCell(cell)
  }

  addNumber(column: number, style: number, value: number): SpreadsheetError | undefined {
    return this.rows.addCell(numberCell(column, value, this.context.styles.format(style), this.context.date1904))
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

/** The RfX range a record starts with (rwFirst, rwLast, colFirst, colLast: u32 each); undefined if truncated or reversed. */
const range = (reader: FixedReader, data: Uint8Array): Range | undefined => {
  const f = reader.fixed(data, RFX_MIN)
  if (f === undefined) return undefined
  const firstRow = f.u32(0)
  const lastRow = f.u32(4)
  const firstColumn = f.u32(8)
  const lastColumn = f.u32(12)
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
    const reader = new FixedReader()
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      if (!scope.step(record)) return undefined
      switch (record.type) {
        case BRT_WS_DIM: {
          const r = range(reader, record.data)
          if (r !== undefined) layout.dimension(r)
          return undefined
        }
        case BRT_COL_INFO: {
          const counted = layout.columnRecord()
          if (counted !== undefined) return counted
          const f = reader.fixed(record.data, COL_INFO_MIN)
          if (f === undefined) return malformed("truncated column info", part, record.offset)
          if ((f.u16(16) & 0x01) !== 0) layout.hideColumns(f.u32(0), f.u32(4))
          return undefined
        }
        case BRT_MERGE_CELL: {
          const r = range(reader, record.data)
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

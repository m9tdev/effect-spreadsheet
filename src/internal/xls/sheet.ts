/**
 * A worksheet substream of a BIFF8/BIFF5 workbook stream -> rows, streamed.
 *
 * The substream starts at its BOUNDSHEET offset with a worksheet BOF and
 * runs to its EOF. Records are read in batches (at most `BATCH_RECORDS`
 * per pull, with a yield before each), so the work between event-loop
 * turns is bounded.
 *
 * - Rows: ROW records give hidden rows (fDyZero) and rows without cells;
 *   cells come row by row after their block of ROW records. A row is
 *   emitted when the cells move on to another row; ROW-only rows are
 *   emitted in order between them. Hidden rows are emitted with
 *   `hidden: true`, never skipped.
 * - Cells: LABELSST, LABEL, RSTRING (text), NUMBER, RK, MULRK (numbers),
 *   BOOLERR (booleans and errors), BLANK, MULBLANK (Empty), FORMULA with
 *   its cached result (a following STRING record for text). Numbers whose
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
import { BOF_BIFF5_8, EOF, type BiffRecord, RecordReader, readBof } from "../biff/records.ts"
import { serialToDate } from "../xlsx/numFmt.ts"
import { decodeRk, ERROR_CODES, Fields, readString } from "./fields.ts"
import type { Globals } from "./globals.ts"

/** Records handled per pull: bounded work between event-loop turns. */
export const BATCH_RECORDS = 4096

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
const STRING = 0x0207
const DIMENSIONS = 0x0200
const COLINFO = 0x007D
const MERGECELLS = 0x00E5

export interface SheetContext {
  readonly stream: Uint8Array
  readonly offset: number
  readonly part: string
  readonly globals: Globals
  readonly limits: Limits
}

const malformed = (part: string, reason: string, offset?: number) =>
  new MalformedFile({ format: "xls", part, reason, ...(offset === undefined ? {} : { offset }) })

/** Opens the substream: checks the offset and the worksheet BOF; the reader after it, or why not. */
const openSubstream = (context: SheetContext): RecordReader | SpreadsheetError => {
  if (context.offset < 0 || context.offset + 4 > context.stream.length) {
    return malformed(context.part, `sheet offset ${context.offset} is outside the workbook stream`)
  }
  const reader = new RecordReader(context.stream, context.part, context.offset)
  const bof = reader.next()
  if (bof === undefined) return malformed(context.part, "the sheet has no BOF", context.offset)
  if (bof instanceof MalformedFile) return bof
  const parsed = readBof(bof)
  if (parsed === undefined) return malformed(context.part, "the sheet doesn't start with a BOF", context.offset)
  return reader
}

/** Builds rows from cell records. */
class Rows {
  out: Array<Row> = []
  private current: number | undefined
  private cells: Array<Cell.Cell> = []
  private sorted = true
  private lastColumn = -1
  /** Hidden flags from ROW records. */
  private readonly hidden = new Set<number>()
  /** Rows declared by ROW records, not emitted yet, in record order. */
  private declared: Array<number> = []
  private declaredAt = 0
  /** Rows emitted so far (a ROW-only row is emitted once). */
  private readonly emitted = new Set<number>()
  private rowCount = 0
  private cellCount = 0
  /** Depth of substreams nested in this one (embedded charts: BOF ... EOF). */
  nested = 0
  /** A FORMULA whose text result comes in the next STRING record. */
  pendingString: { column: number; row: number; meta: Cell.Meta | undefined } | undefined

  private readonly context: SheetContext

  constructor(context: SheetContext) {
    this.context = context
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
      const index = this.declared[this.declaredAt++]!
      if (this.emitted.has(index) || index === this.current) continue
      const error = this.push(index, this.hidden.has(index), [])
      if (error !== undefined) return error
    }
    if (this.declaredAt === this.declared.length) {
      this.declared = []
      this.declaredAt = 0
    }
    return undefined
  }

  /** Ends the current row (if any). */
  private finishRow(): SpreadsheetError | undefined {
    if (this.current === undefined) return undefined
    let cells = this.cells
    if (!this.sorted) {
      // stable by column, keeping the last cell of each column
      const ordered = cells.map((c, k) => [c, k] as const).sort((a, b) => a[0].column - b[0].column || a[1] - b[1])
      cells = []
      // progress: one cell per iteration
      for (let k = 0; k < ordered.length; k++) {
        const next = ordered[k + 1]
        if (next === undefined || next[0].column !== ordered[k]![0].column) cells.push(ordered[k]![0])
      }
    }
    const index = this.current
    this.current = undefined
    this.cells = []
    return this.push(index, this.hidden.has(index), cells)
  }

  row(index: number, hidden: boolean): void {
    if (hidden) this.hidden.add(index)
    this.declared.push(index)
  }

  cell(row: number, cell: Cell.Cell): SpreadsheetError | undefined {
    const limits = this.context.limits
    if (cell.column >= limits.maxColumns) {
      return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: cell.column + 1 })
    }
    if (++this.cellCount > limits.maxCellsPerSheet) {
      return new LimitExceeded({ limit: "maxCellsPerSheet", max: limits.maxCellsPerSheet, actual: this.cellCount })
    }
    if (row !== this.current) {
      const error = this.finishRow() ?? this.flushDeclared(row)
      if (error !== undefined) return error
      this.current = row
      this.sorted = true
      this.lastColumn = -1
    }
    if (cell.column <= this.lastColumn) this.sorted = false
    this.lastColumn = cell.column
    this.cells.push(cell)
    return undefined
  }

  end(): SpreadsheetError | undefined {
    return this.finishRow() ?? this.flushDeclared(Infinity)
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
  return Cell.number(column, value === 0 ? 0 : value, meta)
}

/** Handles one record of the substream; `true` at its EOF. */
const handle = (record: BiffRecord, rows: Rows, context: SheetContext): SpreadsheetError | boolean => {
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
  // a formula's string result must come right after it
  if (rows.pendingString !== undefined && record.type !== STRING) {
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
      if (!Number.isFinite(value)) return malformed(part, "a number that isn't finite", record.offset)
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
        const code = ERROR_CODES.get(value)
        if (code === undefined) return malformed(part, `unknown error code 0x${value.toString(16)}`, record.offset)
        return rows.cell(at[0], Cell.error(at[1], code)) ?? false
      }
      return rows.cell(at[0], Cell.boolean(at[1], value !== 0)) ?? false
    }
    case BLANK: {
      const at = rc()
      if (at === undefined) return truncated()
      return rows.cell(at[0], Cell.empty(at[1])) ?? false
    }
    case FORMULA: {
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
        if (kind === 2) return rows.cell(at[0], Cell.error(at[1], ERROR_CODES.get(value) ?? "#N/A")) ?? false
        return rows.cell(at[0], Cell.text(at[1], "")) ?? false
      }
      const value = f.f64(6)!
      if (!Number.isFinite(value)) return rows.cell(at[0], Cell.error(at[1], "#NUM!")) ?? false
      return rows.cell(at[0], numberCell(globals, at[1], at[2], value)) ?? false
    }
    case STRING: {
      const pending = rows.pendingString
      if (pending === undefined) return false // a STRING without a formula: ignored
      rows.pendingString = undefined
      const text = readString(record.data, 0, globals.text, 2)
      if (text === undefined) return truncated()
      if (text === null) return tooLong()
      return rows.cell(pending.row, Cell.text(pending.column, text.value, pending.meta)) ?? false
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
    const rows = new Rows(context)
    let done = false
    const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, SpreadsheetError> = Effect.gen(function*() {
      if (reader === undefined) return yield* Effect.fail(opened as SpreadsheetError)
      // progress: each iteration reads up to BATCH_RECORDS records (the
      // reader's offset grows with each), or ends the stream
      while (true) {
        if (done) return yield* Cause.done()
        yield* Effect.yieldNow
        let ended = false
        // progress: k up to BATCH_RECORDS; every next() advances the reader
        for (let k = 0; k < BATCH_RECORDS; k++) {
          const record = reader.next()
          if (record === undefined) return yield* Effect.fail(malformed(context.part, "the sheet has no EOF"))
          if (record instanceof MalformedFile) return yield* Effect.fail(record)
          const result = handle(record, rows, context)
          if (result === true) {
            ended = true
            break
          }
          if (result !== false) return yield* Effect.fail(result)
        }
        if (ended) {
          done = true
          const error = rows.end()
          if (error !== undefined) return yield* Effect.fail(error)
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
        if ((grbit & 0x01) !== 0 && last >= first) hidden.push({ first, last: Math.min(last, limits.maxColumns - 1) })
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

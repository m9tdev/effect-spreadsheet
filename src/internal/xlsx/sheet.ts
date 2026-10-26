/**
 * Worksheet part -> rows, streamed: the part is inflated and tokenized
 * piece by piece, and the rows a piece completes are emitted with it (one
 * inflated piece is at most 64 KiB, so the work between emissions is
 * bounded, and the stream yields to the event loop before each piece).
 *
 * - Rows come in document order; `<row r>` gives the index, else the first
 *   cell's reference, else the previous row + 1. Rows without cells are
 *   emitted too; rows the file leaves out are not (rows are sparse).
 *   Hidden rows (`hidden="1"`) are emitted with `hidden: true`.
 * - Cells: `<c r>` gives the column, else the previous cell's + 1. A row
 *   whose cells are out of order is sorted, keeping the last of duplicates.
 * - Values by type (`t`): shared string (`s`), inline string, formula
 *   string (`str`), boolean, error, ISO date (`d`) and number (default).
 *   A number whose style has a date format becomes a `Date` (1900 or 1904
 *   system), unless it is outside Excel's date range. Cells without a value
 *   are `Empty` (with their formula, if any).
 * - Nothing is sized by the declared `<dimension>`; cost follows the cells
 *   present. Limits: row and column indexes, rows and cells per sheet,
 *   string lengths, and the XML limits.
 *
 * @internal
 */
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import * as Cell from "../../Cell.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
import { XmlDecoder } from "../xml/parse.ts"
import { type XmlAttributes, type XmlHandler, XmlTokenizer } from "../xml/tokenizer.ts"
import { serialToDate } from "./numFmt.ts"
import { attr, isTrue, malformed, RichText, type Styles, unescapeXstring } from "./parts.ts"
import { parseCellRef } from "./refs.ts"

export interface SheetContext {
  readonly part: string
  readonly strings: ReadonlyArray<string> | undefined
  readonly styles: Styles
  readonly date1904: boolean
  readonly limits: Limits
}

/** xsd:double as Excel writes it (no hex, no INF/NaN). */
const DOUBLE = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/
/** xsd:dateTime / date, with or without a zone (no zone: taken as UTC wall clock). */
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:\d{2})?)?$/

interface PendingCell {
  column: number
  type: string | undefined
  style: number | undefined
  value: Array<string> | undefined
  valueLength: number
  formula: Array<string> | undefined
  formulaLength: number
  inline: boolean
}

/** Collects rows from tokenizer events. */
class SheetRows implements XmlHandler {
  /** Completed rows not yet emitted. */
  out: Array<Row> = []
  /** Cells emitted so far (for maxCellsPerSheet), and rows. */
  private cells = 0
  private rowCount = 0

  private inSheetData = false
  private rowIndex: number | undefined
  private rowExplicit = false
  private rowHidden = false
  private rowCells: Array<Cell.Cell> = []
  private rowSorted = true
  private lastColumn = -1
  private previousRow = -1

  private cell: PendingCell | undefined
  /** Where cell text goes: the value, the formula, or the inline string. */
  private target: "value" | "formula" | undefined
  private readonly rich: RichText

  private readonly context: SheetContext

  constructor(context: SheetContext) {
    this.context = context
    this.rich = new RichText(context.limits.maxStringLength)
  }

  open = (name: string, attributes: XmlAttributes): SpreadsheetError | undefined => {
    if (this.cell !== undefined) {
      if (this.cell.inline) {
        this.rich.open(name)
        return undefined
      }
      if (name === "v") {
        this.target = "value"
        this.cell.value ??= []
      } else if (name === "f") {
        this.target = "formula"
        this.cell.formula ??= []
      } else if (name === "is") {
        this.cell.inline = true
        this.rich.take()
      }
      return undefined
    }
    if (name === "sheetData") {
      this.inSheetData = true
      return undefined
    }
    if (!this.inSheetData) return undefined
    if (name === "row") return this.openRow(attributes)
    if (name === "c" && this.rowIndex !== undefined) return this.openCell(attributes)
    return undefined
  }

  close = (name: string): SpreadsheetError | undefined => {
    const cell = this.cell
    if (cell !== undefined) {
      if (cell.inline) {
        if (name === "is") {
          cell.inline = false
          cell.value = [this.rich.take()]
          cell.type = "inlineStr"
        } else this.rich.close(name)
        return undefined
      }
      if (name === "v" || name === "f") this.target = undefined
      else if (name === "c") return this.closeCell(cell)
      return undefined
    }
    if (name === "row" && this.rowIndex !== undefined) return this.closeRow()
    if (name === "sheetData") this.inSheetData = false
    return undefined
  }

  text = (value: string): SpreadsheetError | undefined => {
    const cell = this.cell
    if (cell === undefined) return undefined
    if (cell.inline) return this.rich.text(value)
    const max = this.context.limits.maxStringLength
    if (this.target === "value") {
      cell.value!.push(value)
      cell.valueLength += value.length
      if (cell.valueLength > max) return new LimitExceeded({ limit: "maxStringLength", max, actual: cell.valueLength })
    } else if (this.target === "formula") {
      cell.formula!.push(value)
      cell.formulaLength += value.length
      if (cell.formulaLength > max) return new LimitExceeded({ limit: "maxStringLength", max, actual: cell.formulaLength })
    }
    return undefined
  }

  private openRow(attributes: XmlAttributes): SpreadsheetError | undefined {
    const limits = this.context.limits
    const r = attr(attributes, "r")
    if (r !== undefined) {
      if (!/^\d+$/.test(r) || r.length > 10) {
        if (/^\d+$/.test(r)) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: Infinity })
        return malformed(`bad row number ${JSON.stringify(r.slice(0, 20))}`, this.context.part)
      }
      const index = Number(r) - 1
      if (index < 0) return malformed("row number 0", this.context.part)
      if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
      this.rowIndex = index
      this.rowExplicit = true
    } else {
      // decided by the first cell's reference, or previous + 1, at the first cell / row end
      this.rowIndex = -1
      this.rowExplicit = false
    }
    this.rowHidden = isTrue(attr(attributes, "hidden"))
    this.rowCells = []
    this.rowSorted = true
    this.lastColumn = -1
    return undefined
  }

  private openCell(attributes: XmlAttributes): SpreadsheetError | undefined {
    const limits = this.context.limits
    let column = this.lastColumn + 1
    const r = attr(attributes, "r")
    if (r !== undefined) {
      const ref = parseCellRef(r)
      if (ref === undefined || ref.column === undefined) {
        return malformed(`bad cell reference ${JSON.stringify(r.slice(0, 20))}`, this.context.part)
      }
      column = ref.column
      if (!this.rowExplicit && this.rowIndex === -1 && ref.row !== undefined) {
        if (ref.row >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: ref.row + 1 })
        this.rowIndex = ref.row
      }
    }
    if (column >= limits.maxColumns) {
      return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: column + 1 })
    }
    if (column <= this.lastColumn) this.rowSorted = false
    this.lastColumn = column
    const s = attr(attributes, "s")
    this.cell = {
      column,
      type: attr(attributes, "t"),
      style: s !== undefined && /^\d{1,9}$/.test(s) ? Number(s) : undefined,
      value: undefined,
      valueLength: 0,
      formula: undefined,
      formulaLength: 0,
      inline: false
    }
    this.target = undefined
    return undefined
  }

  private closeCell(pending: PendingCell): SpreadsheetError | undefined {
    this.cell = undefined
    this.target = undefined
    const cell = this.toCell(pending)
    if (cell instanceof MalformedFile) return cell
    const max = this.context.limits.maxCellsPerSheet
    if (++this.cells > max) return new LimitExceeded({ limit: "maxCellsPerSheet", max, actual: this.cells })
    this.rowCells.push(cell)
    return undefined
  }

  private toCell(pending: PendingCell): Cell.Cell | MalformedFile {
    const { column } = pending
    const part = this.context.part
    const formulaText = pending.formula === undefined ? "" : pending.formula.join("")
    const meta: Cell.Meta | undefined = formulaText === "" ? undefined : { formula: unescapeXstring(formulaText) }
    if (pending.value === undefined) return Cell.empty(column, meta)
    const raw = pending.value.length === 1 ? pending.value[0]! : pending.value.join("")
    switch (pending.type) {
      case "s": {
        if (raw === "") return Cell.empty(column, meta)
        const strings = this.context.strings
        if (!/^\d{1,9}$/.test(raw.trim())) return malformed(`bad shared string index ${JSON.stringify(raw.slice(0, 20))}`, part)
        const index = Number(raw.trim())
        if (strings === undefined || index >= strings.length) {
          return malformed(`shared string ${index} doesn't exist`, part)
        }
        return Cell.text(column, strings[index]!, meta)
      }
      case "inlineStr":
      case "str":
        return Cell.text(column, pending.type === "str" ? unescapeXstring(raw) : raw, meta)
      case "b": {
        const v = raw.trim()
        if (v === "1" || v === "true") return Cell.boolean(column, true, meta)
        if (v === "0" || v === "false") return Cell.boolean(column, false, meta)
        return malformed(`bad boolean ${JSON.stringify(v.slice(0, 20))}`, part)
      }
      case "e":
        return Cell.error(column, raw.trim(), meta)
      case "d": {
        const v = raw.trim()
        if (v === "") return Cell.empty(column, meta)
        const time = ISO_DATE.test(v) ? Date.parse(/[zZ]|[+-]\d{2}:\d{2}$/.test(v) || !v.includes("T") ? v : `${v}Z`) : NaN
        if (Number.isNaN(time)) return malformed(`bad date ${JSON.stringify(v.slice(0, 40))}`, part)
        return Cell.date(column, new Date(time), this.withFormat(pending, meta))
      }
      case undefined:
      case "n": {
        const v = raw.trim()
        if (v === "") return Cell.empty(column, meta)
        if (!DOUBLE.test(v)) return malformed(`bad number ${JSON.stringify(v.slice(0, 40))}`, part)
        const n = Number(v)
        if (!Number.isFinite(n)) return malformed(`number out of range ${JSON.stringify(v.slice(0, 40))}`, part)
        const format = pending.style === undefined ? undefined : this.context.styles.formats[pending.style]
        const withFormat = this.withFormat(pending, meta)
        if (format?.date === true) {
          const date = serialToDate(n, this.context.date1904)
          if (date !== undefined) return Cell.date(column, date, withFormat)
        }
        return Cell.number(column, n === 0 ? 0 : n, withFormat)
      }
      default:
        return malformed(`unknown cell type ${JSON.stringify(pending.type.slice(0, 20))}`, part)
    }
  }

  private withFormat(pending: PendingCell, meta: Cell.Meta | undefined): Cell.Meta | undefined {
    const code = pending.style === undefined ? undefined : this.context.styles.formats[pending.style]?.code
    if (code === undefined) return meta
    return { ...meta, numFmt: code }
  }

  private closeRow(): SpreadsheetError | undefined {
    const limits = this.context.limits
    let index = this.rowIndex!
    if (index === -1) index = this.previousRow + 1
    if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
    if (++this.rowCount > limits.maxRows) {
      return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: this.rowCount })
    }
    let cells = this.rowCells
    if (!this.rowSorted) {
      // stable sort by column, then keep the last cell of each column
      const sorted = cells.map((c, k) => [c, k] as const).sort((a, b) => a[0].column - b[0].column || a[1] - b[1])
      cells = []
      // progress: one cell per iteration
      for (let k = 0; k < sorted.length; k++) {
        const next = sorted[k + 1]
        if (next === undefined || next[0].column !== sorted[k]![0].column) cells.push(sorted[k]![0])
      }
    }
    this.out.push({ index, hidden: this.rowHidden, cells })
    this.previousRow = index
    this.rowIndex = undefined
    this.rowCells = []
    return undefined
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
      const decoder = new XmlDecoder(context.part)
      const tokenizer = new XmlTokenizer(rows, context.limits, context.part)
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
            const rest = decoder.end()
            error = typeof rest === "string" ? (tokenizer.feed(rest) ?? tokenizer.end()) : rest
          } else {
            // progress: one piece per iteration
            for (const piece of chunk) {
              const text = decoder.push(piece)
              error = typeof text === "string" ? tokenizer.feed(text) : text
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

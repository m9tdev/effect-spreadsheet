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
 * - Row indexes must go up: a row whose index is the same as or below the
 *   previous row's, or a cell whose reference names another row than its
 *   own, is MalformedFile. Rows are streamed, so a row already emitted
 *   can't be merged with a later one.
 * - A `<row>` inside a row, or a `<sheetData>` inside the sheet data, is
 *   MalformedFile: it can't be read without dropping the open row's cells
 *   or the rows after it.
 * - Cells: `<c r>` gives the column, else the previous cell's + 1. A row
 *   whose cells are out of order is sorted, keeping the last of duplicates.
 *   An unsorted row is compacted that way as soon as it holds more than
 *   2 x `maxColumns` cells, so repeated columns don't pile up (a row holds
 *   at most that many cells, and the sorting work is amortized).
 * - Values by type (`t`): shared string (`s`), inline string, formula
 *   string (`str`), boolean, error, ISO date (`d`) and number (default).
 *   A number whose style has a date format becomes a `Date` (1900 or 1904
 *   system), unless it is outside Excel's date range. Cells without a value
 *   are `Empty` (with their formula, if any). The type decides where the
 *   value is: `<is>` for an inline string (else its `<v>`), `<v>` for the
 *   others. Only a cell's first `<v>`, `<f>` and `<is>` count, and only its
 *   direct children: elements nested deeper (in an `<extLst>`) are skipped.
 * - Nothing is sized by the declared `<dimension>`; cost follows the cells
 *   present. Limits: row and column indexes, rows and cells per sheet,
 *   string lengths, and the XML limits.
 * - On an error (malformed input, a limit), the rows completed before it
 *   come out first, then the stream fails with it. So the rows seen before
 *   the error don't depend on where the input is split into pieces.
 *
 * @internal
 */
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import { absurd } from "effect/Function"
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
  /** The text of the cell's first `<v>` (a repeated one is ignored). */
  value: Array<string> | undefined
  valueLength: number
  /** The text of the first `<f>`. */
  formula: Array<string> | undefined
  formulaLength: number
  /** Inside the first `<is>`. */
  inline: boolean
  /** The first `<is>`'s text, unescaped. Used only by an inline string (or a cell without `t` and `<v>`). */
  inlineValue: string | undefined
}

/**
 * Sorts a row's cells by column in place, keeping the last cell of each
 * column (`sort` is stable), and returns them.
 */
export const compact = (cells: Array<Cell.Cell>): Array<Cell.Cell> => {
  cells.sort((a, b) => a.column - b.column)
  let kept = 0
  // progress: k up to cells.length
  for (let k = 0; k < cells.length; k++) {
    const next = cells[k + 1]
    if (next === undefined || next.column !== cells[k]!.column) cells[kept++] = cells[k]!
  }
  cells.length = kept
  return cells
}

/** Collects rows from tokenizer events. Exported for tests. */
export class SheetRows implements XmlHandler {
  /** Completed rows not yet emitted. */
  out: Array<Row> = []
  /** Cells emitted so far (for maxCellsPerSheet), and rows. */
  private cells = 0
  private rowCount = 0

  private inSheetData = false
  /** The open row's index; -1 until its first cell decides it (a row without `r`). */
  private rowIndex: number | undefined
  private rowHidden = false
  private rowCells: Array<Cell.Cell> = []
  /** `rowCells` is in strictly increasing column order. */
  private rowSorted = true
  private lastColumn = -1
  private previousRow = -1

  private cell: PendingCell | undefined
  /** Where cell text goes: the value, the formula, or the inline string. */
  private target: "value" | "formula" | undefined
  /**
   * Elements open inside the cell: 0 between its own children. Only its
   * direct children give its value, and only its own end tag closes it, so
   * elements nested in an extension (`<extLst>`) are skipped, whatever
   * their names (the tokenizer checks that end tags match).
   */
  private depth = 0
  private readonly rich: RichText

  private readonly context: SheetContext

  constructor(context: SheetContext) {
    this.context = context
    this.rich = new RichText(context.limits.maxStringLength)
  }

  open = (name: string, attributes: XmlAttributes): SpreadsheetError | undefined => {
    if (this.cell !== undefined) {
      const depth = this.depth++
      if (this.cell.inline) {
        this.rich.open(name)
        return undefined
      }
      if (depth > 0) return undefined
      // the first <v>, <f> and <is> count; a repeated one is skipped, not appended
      if (name === "v") {
        this.target = this.cell.value === undefined ? "value" : undefined
        this.cell.value ??= []
      } else if (name === "f") {
        this.target = this.cell.formula === undefined ? "formula" : undefined
        this.cell.formula ??= []
      } else if (name === "is" && this.cell.inlineValue === undefined) {
        this.cell.inline = true
        this.rich.take()
      }
      return undefined
    }
    if (name === "sheetData") {
      // a nested one would end the sheet data at its end tag, hiding the rows after it
      if (this.inSheetData) return malformed("<sheetData> inside a <sheetData>", this.context.part)
      this.inSheetData = true
      return undefined
    }
    if (!this.inSheetData) return undefined
    if (name === "row") {
      // a nested one would replace the open row, dropping its cells
      if (this.rowIndex !== undefined) return malformed("<row> inside a <row>", this.context.part)
      return this.openRow(attributes)
    }
    if (name === "c" && this.rowIndex !== undefined) return this.openCell(attributes)
    return undefined
  }

  close = (name: string): SpreadsheetError | undefined => {
    const cell = this.cell
    if (cell !== undefined) {
      const depth = --this.depth
      // below 0: the cell's own </c>
      if (depth < 0) return this.closeCell(cell)
      if (cell.inline) {
        if (depth === 0) {
          // </is>
          cell.inline = false
          cell.inlineValue = this.rich.take()
        } else this.rich.close(name)
        return undefined
      }
      // the end of a direct child: <v>, <f> or another
      if (depth === 0) this.target = undefined
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
    // text directly in the <v> or <f>, not in an element nested in it
    if (this.depth !== 1) return undefined
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
      if (index <= this.previousRow) return this.backwards(index)
      this.rowIndex = index
    } else {
      // decided by the first cell's reference, or previous + 1, at the first cell / row end
      this.rowIndex = -1
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
    let row: number | undefined
    const r = attr(attributes, "r")
    if (r !== undefined) {
      const ref = parseCellRef(r)
      if (ref === undefined || ref.column === undefined) {
        return malformed(`bad cell reference ${JSON.stringify(r.slice(0, 20))}`, this.context.part)
      }
      column = ref.column
      row = ref.row
    }
    if (this.rowIndex === -1) {
      // a row without `r`: its first cell decides, by its reference or as the previous row + 1
      const index = row ?? this.previousRow + 1
      if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
      if (index <= this.previousRow) return this.backwards(index)
      this.rowIndex = index
    } else if (row !== undefined && row !== this.rowIndex) {
      return malformed(`cell ${JSON.stringify(r!.slice(0, 20))} in row ${this.rowIndex! + 1}`, this.context.part)
    }
    if (column >= limits.maxColumns) {
      return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: column + 1 })
    }
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
      inline: false,
      inlineValue: undefined
    }
    this.target = undefined
    this.depth = 0
    return undefined
  }

  private closeCell(pending: PendingCell): SpreadsheetError | undefined {
    this.cell = undefined
    this.target = undefined
    this.depth = 0
    const cell = this.toCell(pending)
    if (cell instanceof MalformedFile) return cell
    const max = this.context.limits.maxCellsPerSheet
    if (++this.cells > max) return new LimitExceeded({ limit: "maxCellsPerSheet", max, actual: this.cells })
    const cells = this.rowCells
    if (cells.length > 0 && cell.column <= cells[cells.length - 1]!.column) this.rowSorted = false
    cells.push(cell)
    // repeated columns don't pile up: past 2 x maxColumns cells an unsorted
    // row is compacted (to at most maxColumns), so the work is amortized
    if (!this.rowSorted && cells.length > 2 * this.context.limits.maxColumns) {
      compact(cells)
      this.rowSorted = true
    }
    return undefined
  }

  /** Cells held for the open row (for tests). */
  get heldCells(): number {
    return this.rowCells.length
  }

  private toCell(pending: PendingCell): Cell.Cell | MalformedFile {
    const { column } = pending
    const part = this.context.part
    const formulaText = pending.formula === undefined ? "" : pending.formula.join("")
    const meta: Cell.Meta | undefined = formulaText === "" ? undefined : { formula: unescapeXstring(formulaText) }
    // the type decides where the value is: <is> for an inline string (else
    // its <v>), <v> for every other type; a cell without `t` or `<v>` takes <is>
    if (pending.inlineValue !== undefined && (pending.type === "inlineStr" || (pending.type === undefined && pending.value === undefined))) {
      return Cell.text(column, pending.inlineValue, meta)
    }
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
        // text in <v> (str, or inlineStr written that way)
        return Cell.text(column, unescapeXstring(raw), meta)
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

  /** A row whose index (0-based) isn't above the previous row's. */
  private backwards(index: number): MalformedFile {
    return malformed(`row ${index + 1} after row ${this.previousRow + 1}`, this.context.part)
  }

  private closeRow(): SpreadsheetError | undefined {
    const limits = this.context.limits
    let index = this.rowIndex!
    if (index === -1) index = this.previousRow + 1
    if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
    if (++this.rowCount > limits.maxRows) {
      return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: this.rowCount })
    }
    const cells = this.rowSorted ? this.rowCells : compact(this.rowCells)
    this.out.push({ index, hidden: this.rowHidden, cells })
    this.previousRow = index
    this.rowIndex = undefined
    this.rowCells = []
    return undefined
  }
}

/**
 * Where a sheet's row stream is: reading input; failing, with the rows
 * completed before the error already emitted; or done (the input ended,
 * or the error was raised). Shared with the XLSB reader.
 */
export type PullState =
  | { readonly _tag: "Reading" }
  | { readonly _tag: "Failing"; readonly error: SpreadsheetError }
  | { readonly _tag: "Done" }

export const READING: PullState = { _tag: "Reading" }
export const DONE: PullState = { _tag: "Done" }

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
      const pullUpstream = Pull.catchDone(upstream, () => Effect.undefined)
      let state: PullState = READING
      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, SpreadsheetError> = Effect.gen(function*() {
        // progress: each iteration pulls one upstream piece (consumed), or
        // ends the input (Done), or ends the stream (Failing goes to Done)
        while (true) {
          switch (state._tag) {
            case "Reading":
              break
            case "Failing": {
              const error = state.error
              state = DONE
              return yield* Effect.fail(error)
            }
            case "Done":
              return yield* Cause.done()
            default:
              return absurd(state)
          }
          yield* Effect.yieldNow
          const chunk = yield* pullUpstream
          let error: SpreadsheetError | undefined
          if (chunk === undefined) {
            state = DONE
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
          // the rows completed before the error come out first; the error on the next pull
          if (error !== undefined) state = { _tag: "Failing", error }
          if (rows.out.length > 0) {
            const out = rows.out
            rows.out = []
            return out as unknown as Arr.NonEmptyReadonlyArray<Row>
          }
        }
      })
      return pull
    }))

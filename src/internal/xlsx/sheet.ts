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
  /** The first `<is>`'s text, unescaped. Used only by an inline string (or a cell without `t` and `<v>`). */
  inlineValue: string | undefined
  /** The direct child the tokenizer is in, which says where text goes. */
  child: CellChild
}

/**
 * Which direct child of a cell is open: none (between its children, or in
 * one that doesn't count: a repeated `<v>`/`<f>`/`<is>`, or another element),
 * its first `<v>` (text goes to the value), its first `<f>` (to the formula),
 * or its first `<is>` (to the inline string).
 */
type CellChild = "none" | "value" | "formula" | "inline"

/** The open row's fields, whether its index is known yet or not. */
interface RowFields {
  readonly hidden: boolean
  readonly cells: Array<Cell.Cell>
  /** `cells` is in strictly increasing column order. */
  sorted: boolean
  lastColumn: number
}

/**
 * The open row: numbered (by its `r`, or by its first cell), or unnumbered
 * (a row without `r` before its first cell, which decides the index; a row
 * without `r` or cells is the previous row + 1).
 */
type OpenRow =
  | RowFields & { readonly _tag: "Numbered"; readonly index: number }
  | RowFields & { readonly _tag: "Unnumbered"; readonly index: undefined }

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
  /** The open row, if any (only inside the sheet data). */
  private row: OpenRow | undefined
  private previousRow = -1

  /** The open cell, if any (only inside a row). */
  private cell: PendingCell | undefined
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
    const cell = this.cell
    if (cell !== undefined) {
      const depth = this.depth++
      if (cell.child === "inline") {
        this.rich.open(name)
        return undefined
      }
      // a direct child (between children, no child is open)
      if (depth === 0) cell.child = this.openChild(cell, name)
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
      if (this.row !== undefined) return malformed("<row> inside a <row>", this.context.part)
      return this.openRow(attributes)
    }
    if (name === "c" && this.row !== undefined) return this.openCell(this.row, attributes)
    return undefined
  }

  close = (name: string): SpreadsheetError | undefined => {
    const cell = this.cell
    if (cell !== undefined) {
      const depth = --this.depth
      // below 0: the cell's own </c>
      if (depth < 0) return this.closeCell(cell)
      if (depth > 0) {
        if (cell.child === "inline") this.rich.close(name)
        return undefined
      }
      // the end of a direct child: <v>, <f>, <is> or another
      if (cell.child === "inline") cell.inlineValue = this.rich.take()
      cell.child = "none"
      return undefined
    }
    if (name === "row" && this.row !== undefined) return this.closeRow(this.row)
    if (name === "sheetData") this.inSheetData = false
    return undefined
  }

  text = (value: string): SpreadsheetError | undefined => {
    const cell = this.cell
    if (cell === undefined) return undefined
    switch (cell.child) {
      case "none":
        return undefined
      case "inline":
        return this.rich.text(value)
      case "value": {
        // text directly in the <v>, not in an element nested in it
        if (this.depth !== 1) return undefined
        cell.value!.push(value)
        cell.valueLength += value.length
        const max = this.context.limits.maxStringLength
        if (cell.valueLength > max) return new LimitExceeded({ limit: "maxStringLength", max, actual: cell.valueLength })
        return undefined
      }
      case "formula": {
        // text directly in the <f>, not in an element nested in it
        if (this.depth !== 1) return undefined
        cell.formula!.push(value)
        cell.formulaLength += value.length
        const max = this.context.limits.maxStringLength
        if (cell.formulaLength > max) return new LimitExceeded({ limit: "maxStringLength", max, actual: cell.formulaLength })
        return undefined
      }
      default:
        return absurd(cell.child)
    }
  }

  /**
   * The child a direct child element of the cell opens. Only the first
   * `<v>`, `<f>` and `<is>` count; a repeated one is skipped, not appended.
   */
  private openChild(cell: PendingCell, name: string): CellChild {
    if (name === "v") {
      if (cell.value !== undefined) return "none"
      cell.value = []
      return "value"
    }
    if (name === "f") {
      if (cell.formula !== undefined) return "none"
      cell.formula = []
      return "formula"
    }
    if (name === "is" && cell.inlineValue === undefined) {
      this.rich.take()
      return "inline"
    }
    return "none"
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
      this.row = { _tag: "Numbered", index, hidden: isTrue(attr(attributes, "hidden")), cells: [], sorted: true, lastColumn: -1 }
    } else {
      // decided by the first cell's reference, or previous + 1, at the first cell / row end
      this.row = { _tag: "Unnumbered", index: undefined, hidden: isTrue(attr(attributes, "hidden")), cells: [], sorted: true, lastColumn: -1 }
    }
    return undefined
  }

  private openCell(open: OpenRow, attributes: XmlAttributes): SpreadsheetError | undefined {
    const limits = this.context.limits
    let column = open.lastColumn + 1
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
    switch (open._tag) {
      case "Unnumbered": {
        // a row without `r`: its first cell decides, by its reference or as the previous row + 1
        const index = row ?? this.previousRow + 1
        if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
        if (index <= this.previousRow) return this.backwards(index)
        open = { _tag: "Numbered", index, hidden: open.hidden, cells: open.cells, sorted: open.sorted, lastColumn: open.lastColumn }
        this.row = open
        break
      }
      case "Numbered":
        if (row !== undefined && row !== open.index) {
          return malformed(`cell ${JSON.stringify(r!.slice(0, 20))} in row ${open.index + 1}`, this.context.part)
        }
        break
      default:
        return absurd(open)
    }
    if (column >= limits.maxColumns) {
      return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: column + 1 })
    }
    open.lastColumn = column
    const s = attr(attributes, "s")
    this.cell = {
      column,
      type: attr(attributes, "t"),
      style: s !== undefined && /^\d{1,9}$/.test(s) ? Number(s) : undefined,
      value: undefined,
      valueLength: 0,
      formula: undefined,
      formulaLength: 0,
      inlineValue: undefined,
      child: "none"
    }
    this.depth = 0
    return undefined
  }

  private closeCell(pending: PendingCell): SpreadsheetError | undefined {
    this.cell = undefined
    this.depth = 0
    // a cell is only open inside a row (opened by its first cell, if it had no `r`)
    const open = this.row!
    const cell = this.toCell(pending)
    if (cell instanceof MalformedFile) return cell
    const max = this.context.limits.maxCellsPerSheet
    if (++this.cells > max) return new LimitExceeded({ limit: "maxCellsPerSheet", max, actual: this.cells })
    const cells = open.cells
    if (cells.length > 0 && cell.column <= cells[cells.length - 1]!.column) open.sorted = false
    cells.push(cell)
    // repeated columns don't pile up: past 2 x maxColumns cells an unsorted
    // row is compacted (to at most maxColumns), so the work is amortized
    if (!open.sorted && cells.length > 2 * this.context.limits.maxColumns) {
      compact(cells)
      open.sorted = true
    }
    return undefined
  }

  /** Cells held for the open row (for tests). */
  get heldCells(): number {
    return this.row === undefined ? 0 : this.row.cells.length
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

  private closeRow(open: OpenRow): SpreadsheetError | undefined {
    const limits = this.context.limits
    let index: number
    switch (open._tag) {
      case "Numbered":
        index = open.index
        break
      case "Unnumbered":
        // a row without `r` or cells
        index = this.previousRow + 1
        break
      default:
        return absurd(open)
    }
    if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
    if (++this.rowCount > limits.maxRows) {
      return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: this.rowCount })
    }
    const cells = open.sorted ? open.cells : compact(open.cells)
    this.out.push({ index, hidden: open.hidden, cells })
    this.previousRow = index
    this.row = undefined
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

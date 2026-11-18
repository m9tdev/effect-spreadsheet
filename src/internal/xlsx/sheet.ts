/**
 * Worksheet part -> rows, streamed: the part is inflated and tokenized
 * piece by piece, and the rows a piece completes are emitted with it. The
 * archive hands the part over in pieces of at most 64 KiB, each after a
 * yield to the event loop, so the work between emissions is bounded.
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
import type { Stream } from "effect"
import { absurd } from "effect/Function"
import * as Cell from "../../Cell.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
import { decoded, step } from "../decode.ts"
import { type CellStyles, numberCell, type NumberFormat, withNumFmt } from "../numFmt.ts"
import { parseCellRef } from "../refs.ts"
import { RowBuilder } from "../sheet/rows.ts"
import { xmlDecoder } from "../xml/parse.ts"
import { attr, type XmlAttributes, type XmlHandler } from "../xml/tokenizer.ts"
import { isTrue, malformed, RichText, unescapeXstring } from "./parts.ts"

export interface SheetContext {
  readonly part: string
  readonly strings: ReadonlyArray<string> | undefined
  readonly styles: CellStyles
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

/**
 * Where a row or cell sits, for errors: a cell's `r` text, if it has one.
 * XLSX errors name the part, not an offset.
 */
type At = string | undefined

/** Collects rows from tokenizer events. Exported for tests. */
export class SheetRows implements XmlHandler {
  /**
   * The rows (only inside the sheet data): in XLSX a cell outside a row is
   * ignored, and a row without `r` is numbered by its first cell.
   */
  readonly rows: RowBuilder<At>

  /** The completed rows not yet taken. */
  take(): ReadonlyArray<Row> {
    return this.rows.take()
  }

  private inSheetData = false

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
    const part = context.part
    this.rows = new RowBuilder<At>(context.limits, {
      cellOutsideRow: "ignore",
      malformed: (reason) => malformed(reason, part),
      cellInOtherRow: (row, ref) => malformed(`cell ${JSON.stringify((ref ?? "").slice(0, 20))} in row ${row + 1}`, part)
    })
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
      if (this.rows.inRow) return malformed("<row> inside a <row>", this.context.part)
      return this.openRow(attributes)
    }
    if (name === "c") {
      const start = this.rows.beginCell(undefined)
      if (start !== "read") return start === "skip" ? undefined : start
      return this.openCell(attributes)
    }
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
    if (name === "row" && this.rows.inRow) return this.rows.closeRow()
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

  /** A `<row>`: its `r` gives the index; without one, its first cell or the previous row decides. */
  private openRow(attributes: XmlAttributes): SpreadsheetError | undefined {
    const hidden = isTrue(attr(attributes, "hidden"))
    const r = attr(attributes, "r")
    if (r === undefined) return this.rows.openRow(undefined, hidden, undefined)
    if (!/^\d+$/.test(r) || r.length > 10) {
      if (/^\d+$/.test(r)) return new LimitExceeded({ limit: "maxRows", max: this.context.limits.maxRows, actual: Infinity })
      return malformed(`bad row number ${JSON.stringify(r.slice(0, 20))}`, this.context.part)
    }
    const index = Number(r) - 1
    if (index < 0) return malformed("row number 0", this.context.part)
    return this.rows.openRow(index, hidden, undefined)
  }

  /** A `<c>` in a row: its `r` gives the column (and the row), else the previous cell's column + 1. */
  private openCell(attributes: XmlAttributes): SpreadsheetError | undefined {
    const rows = this.rows
    let column = rows.nextColumn
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
    const placed = rows.placeCell(column, r, row)
    if (placed !== undefined) return placed
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
    const cell = this.toCell(pending)
    if (cell instanceof MalformedFile) return cell
    return this.rows.addCell(cell)
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
        return Cell.date(column, new Date(time), withNumFmt(meta, this.format(pending)))
      }
      case undefined:
      case "n": {
        const v = raw.trim()
        if (v === "") return Cell.empty(column, meta)
        if (!DOUBLE.test(v)) return malformed(`bad number ${JSON.stringify(v.slice(0, 40))}`, part)
        const n = Number(v)
        // XLSX's own rule: text out of the double range is a malformed file (not #NUM!)
        if (!Number.isFinite(n)) return malformed(`number out of range ${JSON.stringify(v.slice(0, 40))}`, part)
        return numberCell(column, n, this.format(pending), this.context.date1904, meta)
      }
      default:
        return malformed(`unknown cell type ${JSON.stringify(pending.type.slice(0, 20))}`, part)
    }
  }

  private format(pending: PendingCell): NumberFormat | undefined {
    return pending.style === undefined ? undefined : this.context.styles.format(pending.style)
  }
}

/**
 * Rows of a worksheet part, given its byte stream. Re-runnable when the
 * input is. The tokenizer hands its work to the work meter.
 */
export const sheetRows = (
  bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
  context: SheetContext
): Stream.Stream<Row, SpreadsheetError> =>
  decoded(bytes, (meter) => {
    const rows = new SheetRows(context)
    const xml = xmlDecoder(rows, context.limits, context.part, meter)
    return {
      feed: (piece) => {
        const error = xml.feed(piece).error
        return step(rows.take(), error)
      },
      end: () => {
        const error = xml.end().error
        return step(rows.take(), error)
      }
    }
  })

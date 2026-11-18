/**
 * The format-neutral row builder: the one place the `Row` contract and the
 * per-sheet row and cell limits live. A format's sheet reader decodes its
 * records or elements and hands the builder rows and cells; the builder
 * emits rows that keep the contract:
 *
 * - row indexes go up strictly, each below `maxRows`, and at most
 *   `maxRows` rows are emitted;
 * - a row's cells are in ascending column order, one cell per column (the
 *   last cell given for a column wins), each column below `maxColumns`;
 * - at most `maxCellsPerSheet` cells are read, repeats included;
 * - hidden rows are emitted with `hidden: true`, never skipped; rows
 *   without cells are emitted too.
 *
 * Cells are held only for the open row. A row whose cells come out of
 * order is sorted (and compacted) once, when it closes, or as soon as it
 * holds more than 2 x `maxColumns` cells, so the cells held stay under that
 * however often the file repeats a column.
 *
 * Where the formats differ, the rule is a parameter (`RowRules`), never
 * unified here: what a cell outside a row is, and how a malformed row
 * order is reported.
 *
 * @internal
 */
import type * as Cell from "../../Cell.ts"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
import { compact, NO_ROWS } from "../sheet.ts"

/** A format's own rules. `At` is where a row or cell sits in the source (a record, say), for errors. */
export interface RowRules<At> {
  /** A cell with no open row: skipped (`beginCell` gives `IGNORED`), or this error. */
  readonly cellOutsideRow: "ignore" | ((at: At) => SpreadsheetError)
  /** A MalformedFile with this reason, at `at`. */
  readonly malformed: (reason: string, at: At) => SpreadsheetError
}

/** What `beginCell` gives for a cell outside a row under the "ignore" rule. */
export const IGNORED: unique symbol = Symbol("IGNORED")

/** Builds a sheet's rows from its rows and cells, in document order. */
export class RowBuilder<At> {
  /** Completed rows not yet taken. */
  private out: Array<Row> = []
  /** Cells read so far, repeats included (for maxCellsPerSheet), and rows opened (for maxRows). */
  private cellCount = 0
  private rowCount = 0
  /** The last row's index, -1 before the first row. */
  private previousRow = -1

  /** Whether a row is open; the fields below are the open row's. */
  private open = false
  private index = 0
  private hidden = false
  private cells: Array<Cell.Cell> = []
  /** `cells` is in strictly increasing column order. */
  private sorted = true
  /** The last cell's column, -1 before the first. */
  private lastColumn = -1

  private readonly limits: Limits
  private readonly rules: RowRules<At>

  constructor(limits: Limits, rules: RowRules<At>) {
    this.limits = limits
    this.rules = rules
  }

  /** The completed rows not yet taken; the builder starts collecting anew. */
  take(): ReadonlyArray<Row> {
    const out = this.out
    if (out.length === 0) return NO_ROWS
    this.out = []
    return out
  }

  /** Whether a row is open. */
  get inRow(): boolean {
    return this.open
  }

  /** Cells held for the open row (0 when none is open). */
  get held(): number {
    return this.open ? this.cells.length : 0
  }

  /** The column a cell without its own takes: the previous cell's in the row + 1 (0 for the row's first). */
  get nextColumn(): number {
    return this.lastColumn + 1
  }

  /**
   * Opens row `index` (0-based), closing the open row first, if any. Fails
   * on an index at or past `maxRows`, or one not above the previous row's.
   */
  openRow(index: number, hidden: boolean, at: At): SpreadsheetError | undefined {
    const closed = this.closeRow()
    if (closed !== undefined) return closed
    const limits = this.limits
    if (index >= limits.maxRows) return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: index + 1 })
    if (index <= this.previousRow) return this.rules.malformed(`row ${index + 1} after row ${this.previousRow + 1}`, at)
    // a guard: whole-number indexes go up and stay below maxRows, so they
    // never number more rows than that; only a fractional index reaches it
    if (++this.rowCount > limits.maxRows) {
      return new LimitExceeded({ limit: "maxRows", max: limits.maxRows, actual: this.rowCount })
    }
    this.previousRow = index
    this.open = true
    this.index = index
    this.hidden = hidden
    this.cells = []
    this.sorted = true
    this.lastColumn = -1
    return undefined
  }

  /**
   * Closes the open row, if any: it is complete, and goes out with its cells
   * sorted and compacted.
   */
  closeRow(): SpreadsheetError | undefined {
    if (!this.open) return undefined
    this.out.push({ index: this.index, hidden: this.hidden, cells: this.sorted ? this.cells : compact(this.cells) })
    this.open = false
    return undefined
  }

  /** A cell starts, at `at`: `undefined` inside a row; outside one, `IGNORED` or the error, per the format's rule. */
  beginCell(at: At): SpreadsheetError | typeof IGNORED | undefined {
    if (this.open) return undefined
    const outside = this.rules.cellOutsideRow
    return outside === "ignore" ? IGNORED : outside(at)
  }

  /**
   * The cell begun is in `column`: fails at or past `maxColumns`. Sets the
   * row's last column (`nextColumn`) before the cell's value is decoded.
   */
  placeCell(column: number): SpreadsheetError | undefined {
    const max = this.limits.maxColumns
    if (column >= max) return new LimitExceeded({ limit: "maxColumns", max, actual: column + 1 })
    this.lastColumn = column
    return undefined
  }

  /** The cell begun and placed, decoded: added to the open row. Fails past `maxCellsPerSheet` cells. */
  addCell(cell: Cell.Cell): SpreadsheetError | undefined {
    const max = this.limits.maxCellsPerSheet
    if (++this.cellCount > max) return new LimitExceeded({ limit: "maxCellsPerSheet", max, actual: this.cellCount })
    const cells = this.cells
    if (cells.length > 0 && cell.column <= cells[cells.length - 1]!.column) this.sorted = false
    cells.push(cell)
    // repeated columns don't pile up: past 2 x maxColumns cells an unsorted
    // row is compacted (to at most maxColumns), so the work is amortized
    if (!this.sorted && cells.length > 2 * this.limits.maxColumns) {
      compact(cells)
      this.sorted = true
    }
    return undefined
  }
}

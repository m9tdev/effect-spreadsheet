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
 * order or a cell naming another row is reported.
 *
 * A row is numbered when it opens (`openRow(index)`), or, opened without an
 * index (`openRow(undefined)`), by its first cell: by the row the cell names
 * (`placeCell(column, at, rowOfRef)`), else as the previous row + 1; a row
 * without an index or cells is the previous row + 1 when it closes. Once
 * numbered, a cell that names another row is the format's error.
 *
 * @internal
 */
import type * as Cell from "../../Cell.ts"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"

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

/** No rows: what a step that completed none emits. */
export const NO_ROWS: ReadonlyArray<Row> = []

/** A format's own rules. `At` is where a row or cell sits in the source (a record, say), for errors. */
export interface RowRules<At> {
  /** A cell with no open row: skipped (`beginCell` gives "skip"), or this error. */
  readonly cellOutsideRow: "ignore" | ((at: At) => SpreadsheetError)
  /** A MalformedFile with this reason, at `at`. */
  readonly malformed: (reason: string, at: At) => SpreadsheetError
  /**
   * A cell at `at` whose own reference names another row than the open one,
   * `row` (0-based). Only formats whose cells name their row (XLSX) reach it.
   */
  readonly cellInOtherRow: (row: number, at: At) => SpreadsheetError
}

/**
 * What `beginCell` says about a cell: "read" it (a row is open), "skip" it
 * (it is outside a row, and the format ignores those), or fail with the
 * format's error.
 */
export type CellStart = "read" | "skip" | SpreadsheetError

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
  /** The open row's index, undefined until it is numbered (opened without one, before its first cell). */
  private index: number | undefined
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
   * Without an index, the row is numbered by its first cell (`placeCell`),
   * or when it closes.
   */
  openRow(index: number | undefined, hidden: boolean, at: At): SpreadsheetError | undefined {
    const closed = this.closeRow()
    if (closed !== undefined) return closed
    if (index !== undefined) {
      const numbered = this.number(index, at)
      if (numbered !== undefined) return numbered
    }
    this.open = true
    this.index = index
    this.hidden = hidden
    this.cells = []
    this.sorted = true
    this.lastColumn = -1
    return undefined
  }

  /** Numbers the open row (or the one opening) `index`: it must be below `maxRows` and above the previous row. */
  private number(index: number, at: At): SpreadsheetError | undefined {
    const tooHigh = this.belowMaxRows(index)
    if (tooHigh !== undefined) return tooHigh
    if (index <= this.previousRow) return this.rules.malformed(`row ${index + 1} after row ${this.previousRow + 1}`, at)
    return this.counted(index)
  }

  private belowMaxRows(index: number): SpreadsheetError | undefined {
    const max = this.limits.maxRows
    return index >= max ? new LimitExceeded({ limit: "maxRows", max, actual: index + 1 }) : undefined
  }

  /** Counts the row numbered `index` and makes it the previous row. */
  private counted(index: number): SpreadsheetError | undefined {
    // a guard: whole-number indexes go up and stay below maxRows, so they
    // never number more rows than that; only a fractional index reaches it
    const max = this.limits.maxRows
    if (++this.rowCount > max) return new LimitExceeded({ limit: "maxRows", max, actual: this.rowCount })
    this.previousRow = index
    this.index = index
    return undefined
  }

  /**
   * Closes the open row, if any: it is complete, and goes out with its cells
   * sorted and compacted.
   */
  closeRow(): SpreadsheetError | undefined {
    if (!this.open) return undefined
    let index = this.index
    if (index === undefined) {
      // a row without an index or cells: the previous row + 1, above it by construction
      index = this.previousRow + 1
      const numbered = this.belowMaxRows(index) ?? this.counted(index)
      if (numbered !== undefined) return numbered
    }
    this.out.push({ index, hidden: this.hidden, cells: this.sorted ? this.cells : compact(this.cells) })
    this.open = false
    return undefined
  }

  /** A cell starts, at `at`: "read" inside a row; outside one, "skip" or the error, per the format's rule. */
  beginCell(at: At): CellStart {
    if (this.open) return "read"
    const outside = this.rules.cellOutsideRow
    return outside === "ignore" ? "skip" : outside(at)
  }

  /**
   * The cell begun (and read) is in `column`, and, if its reference names
   * one, in row `rowOfRef`. The row's first cell numbers a row opened
   * without an index: `rowOfRef`, else the previous row + 1. Fails on a
   * `rowOfRef` other than the open row's, and at or past `maxColumns`. Sets
   * the row's last column (which `nextColumn` follows) before the cell's
   * value is decoded.
   */
  placeCell(column: number, at: At, rowOfRef?: number): SpreadsheetError | undefined {
    const index = this.index
    if (index === undefined) {
      const numbered = this.number(rowOfRef ?? this.previousRow + 1, at)
      if (numbered !== undefined) return numbered
    } else if (rowOfRef !== undefined && rowOfRef !== index) {
      return this.rules.cellInOtherRow(index, at)
    }
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

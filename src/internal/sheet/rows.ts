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
 * however often the file repeats a column (`RowCells`).
 *
 * A format whose rows don't come one after another (XLS: a row's cells may
 * be split, or come after a later row's) uses `HeldRows` instead: it holds
 * each row's cells until the format finishes the row, with the same limits
 * and compaction. The format finishes rows in increasing row order.
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

/**
 * One row's cells as a format gives them: in any column order, a column
 * maybe repeated. They are sorted (and compacted) once, when the row is
 * done, or as soon as they are more than 2 x `maxColumns` and out of order,
 * so the cells held stay under that however often the file repeats a column.
 */
export class RowCells {
  /** The cells so far. */
  cells: Array<Cell.Cell>
  /** `cells` is in strictly increasing column order. */
  sorted = true
  /** The last cell's column, -1 before the first. */
  private last = -1

  /**
   * Starts with `cells`, empty. The caller allocates the array: V8 keeps
   * the array's allocation feedback at the caller's `[]`, and an array
   * allocated here instead was measurably slower to fill (a wide row).
   */
  constructor(cells: Array<Cell.Cell>) {
    this.cells = cells
  }

  /** Adds `cell`; returns how many cells it sorted, compacting the row (0 when it didn't). */
  add(cell: Cell.Cell, maxColumns: number): number {
    const cells = this.cells
    const column = cell.column
    if (column <= this.last) this.sorted = false
    this.last = column
    cells.push(cell)
    // repeated columns don't pile up: past 2 x maxColumns cells an unsorted
    // row is compacted (to at most maxColumns), so the work is amortized
    if (this.sorted || cells.length <= 2 * maxColumns) return 0
    const sorting = cells.length
    compact(cells)
    this.sorted = true
    this.last = cells[cells.length - 1]!.column
    return sorting
  }

  /** The cells, sorted and compacted. */
  done(): Array<Cell.Cell> {
    return this.sorted ? this.cells : compact(this.cells)
  }

  /** Starts a new row with `cells`, empty (see the constructor); the cells given out by `done` stay as they are. */
  reset(cells: Array<Cell.Cell>): void {
    this.cells = cells
    this.sorted = true
    this.last = -1
  }
}

/** The per-sheet limits on rows, columns and cells, with the counts they need. */
class SheetCounter {
  /** Cells read so far, repeats included (for maxCellsPerSheet), and rows numbered (for maxRows). */
  cells = 0
  rows = 0
  readonly limits: Limits

  constructor(limits: Limits) {
    this.limits = limits
  }

  /** The maxColumns error, if `column` is at or past it. */
  column(column: number): SpreadsheetError | undefined {
    const max = this.limits.maxColumns
    return column >= max ? new LimitExceeded({ limit: "maxColumns", max, actual: column + 1 }) : undefined
  }

  /** Counts a cell read: fails past `maxCellsPerSheet` cells. */
  cell(): SpreadsheetError | undefined {
    const max = this.limits.maxCellsPerSheet
    return ++this.cells > max ? new LimitExceeded({ limit: "maxCellsPerSheet", max, actual: this.cells }) : undefined
  }

  /** The maxRows error, if row `index` is at or past `maxRows`. */
  rowIndex(index: number): SpreadsheetError | undefined {
    const max = this.limits.maxRows
    return index >= max ? new LimitExceeded({ limit: "maxRows", max, actual: index + 1 }) : undefined
  }

  /** Counts a row numbered. */
  row(): SpreadsheetError | undefined {
    // a guard: whole-number indexes go up and stay below maxRows, so they
    // never number more rows than that; only a fractional index reaches it
    const max = this.limits.maxRows
    return ++this.rows > max ? new LimitExceeded({ limit: "maxRows", max, actual: this.rows }) : undefined
  }
}

/** Builds a sheet's rows from its rows and cells, in document order. */
export class RowBuilder<At> {
  /** Completed rows not yet taken. */
  private out: Array<Row> = []
  private readonly count: SheetCounter
  /** The last row's index, -1 before the first row. */
  private previousRow = -1

  /** Whether a row is open; the fields below are the open row's. */
  private open = false
  /** The open row's index, undefined until it is numbered (opened without one, before its first cell). */
  private index: number | undefined
  private hidden = false
  private readonly row = new RowCells([])
  /** The last cell's column, -1 before the first. */
  private lastColumn = -1

  private readonly rules: RowRules<At>

  constructor(limits: Limits, rules: RowRules<At>) {
    this.count = new SheetCounter(limits)
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
    return this.open ? this.row.cells.length : 0
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
    this.row.reset([])
    this.lastColumn = -1
    return undefined
  }

  /** Numbers the open row (or the one opening) `index`: it must be below `maxRows` and above the previous row. */
  private number(index: number, at: At): SpreadsheetError | undefined {
    const tooHigh = this.count.rowIndex(index)
    if (tooHigh !== undefined) return tooHigh
    if (index <= this.previousRow) return this.rules.malformed(`row ${index + 1} after row ${this.previousRow + 1}`, at)
    return this.counted(index)
  }

  /** Counts the row numbered `index` and makes it the previous row. */
  private counted(index: number): SpreadsheetError | undefined {
    const counted = this.count.row()
    if (counted !== undefined) return counted
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
      const numbered = this.count.rowIndex(index) ?? this.counted(index)
      if (numbered !== undefined) return numbered
    }
    this.out.push({ index, hidden: this.hidden, cells: this.row.done() })
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
    const tooFar = this.count.column(column)
    if (tooFar !== undefined) return tooFar
    this.lastColumn = column
    return undefined
  }

  /** The cell begun and placed, decoded: added to the open row. Fails past `maxCellsPerSheet` cells. */
  addCell(cell: Cell.Cell): SpreadsheetError | undefined {
    const counted = this.count.cell()
    if (counted !== undefined) return counted
    this.row.add(cell, this.count.limits.maxColumns)
    return undefined
  }
}

/**
 * Builds a sheet's rows from cells that come in any row order, for formats
 * whose writers may split a row's cells or put rows out of order (XLS).
 * Each row's cells are held until the format finishes the row, which it
 * does in increasing row order. The limits and the per-row compaction are
 * the `RowBuilder`'s: columns below `maxColumns` and at most
 * `maxCellsPerSheet` cells, checked as each cell is added; rows below
 * `maxRows`, checked as each is finished.
 */
export class HeldRows {
  /** Finished rows not yet taken. */
  private out: Array<Row> = []
  /** Rows with cells, not finished yet. */
  private readonly rows = new Map<number, RowCells>()
  /** The row the last cell went to (most cells follow the one before). */
  private hotIndex = -1
  private hot: RowCells | undefined
  private readonly count: SheetCounter
  // the counts below are fields to read, never to set: a format reads them
  // once per record, also before the code is optimized, where a getter
  // call costs more than the read
  /** Cells held: those of the rows not finished yet. */
  held = 0
  /** Cells sorted so far, compacting rows: the sorting work. */
  sorted = 0
  /** The work done so far: cells added, cells sorted and rows finished. */
  work = 0

  constructor(limits: Limits) {
    this.count = new SheetCounter(limits)
  }

  /** The finished rows not yet taken; the collector starts collecting anew. */
  take(): ReadonlyArray<Row> {
    const out = this.out
    if (out.length === 0) return NO_ROWS
    this.out = []
    return out
  }

  /** Whether finished rows wait to be taken. */
  get ready(): boolean {
    return this.out.length > 0
  }

  /** Whether row `index` holds cells. */
  has(index: number): boolean {
    return this.rows.has(index)
  }

  /** The rows holding cells, in row order. */
  indexes(): Uint32Array {
    return Uint32Array.from(this.rows.keys()).sort()
  }

  /** Adds `cell` to row `index`. Fails at or past `maxColumns`, and past `maxCellsPerSheet` cells. */
  addCell(index: number, cell: Cell.Cell): SpreadsheetError | undefined {
    const count = this.count
    const failed = count.column(cell.column) ?? count.cell()
    if (failed !== undefined) return failed
    let row = this.hotIndex === index ? this.hot : this.rows.get(index)
    if (row === undefined) {
      row = new RowCells([])
      this.rows.set(index, row)
    }
    this.hotIndex = index
    this.hot = row
    const before = row.cells.length
    const sorted = row.add(cell, count.limits.maxColumns)
    this.held += row.cells.length - before
    this.sorted += sorted
    this.work += 1 + sorted
    return undefined
  }

  /**
   * Finishes row `index`: it goes out with its cells, sorted and compacted,
   * or without any if it holds none. Fails on an index at or past `maxRows`.
   */
  finishRow(index: number, hidden: boolean): SpreadsheetError | undefined {
    const row = this.rows.get(index)
    if (row !== undefined) {
      this.rows.delete(index)
      this.held -= row.cells.length
      if (this.hotIndex === index) {
        this.hotIndex = -1
        this.hot = undefined
      }
      if (!row.sorted) {
        this.sorted += row.cells.length
        this.work += row.cells.length
      }
    }
    const failed = this.count.rowIndex(index) ?? this.count.row()
    if (failed !== undefined) return failed
    this.work++
    this.out.push({ index, hidden, cells: row === undefined ? [] : row.done() })
    return undefined
  }
}

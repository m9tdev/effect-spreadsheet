import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { LayoutBuilder } from "../src/internal/sheet/layout.ts"
import { HeldRows, type RowRules, RowBuilder } from "../src/internal/sheet/rows.ts"
import { MalformedFile } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"

const malformed = (reason: string, at: number) => new MalformedFile({ format: "xlsx", reason, offset: at })
const rules = (overrides: Partial<RowRules<number>> = {}): RowRules<number> => ({
  cellOutsideRow: (at) => malformed("cell outside a row", at),
  malformed,
  cellInOtherRow: (row, at) => malformed(`cell in row ${row + 1} names another row`, at),
  ...overrides
})
const limits = (overrides: Partial<SpreadsheetLimits.Limits> = {}): SpreadsheetLimits.Limits => ({ ...SpreadsheetLimits.defaults, ...overrides })

/** Adds a cell to the open row, as a sheet reader does. */
const add = (rows: RowBuilder<number>, cell: Cell.Cell) => {
  expect(rows.beginCell(0)).toBe("read")
  return rows.placeCell(cell.column, 0) ?? rows.addCell(cell)
}
const shape = (rows: RowBuilder<number>) => rows.take().map((r) => [r.index, r.hidden, r.cells.map((c) => [c.column, Cell.value(c)])])

describe("row builder", () => {
  it("emits rows in order with their hidden flags, cells sorted with the last of a column kept, and empty rows", () => {
    const rows = new RowBuilder(limits(), rules())
    expect(rows.openRow(0, false, 0)).toBeUndefined()
    expect(add(rows, Cell.number(2, 1))).toBeUndefined()
    expect(add(rows, Cell.number(0, 2))).toBeUndefined()
    expect(add(rows, Cell.number(2, 3))).toBeUndefined()
    expect(rows.openRow(3, true, 1)).toBeUndefined()
    expect(rows.openRow(5, false, 2)).toBeUndefined()
    expect(add(rows, Cell.number(1, 4))).toBeUndefined()
    expect(rows.closeRow()).toBeUndefined()
    expect(rows.closeRow()).toBeUndefined()
    expect(shape(rows)).toEqual([[0, false, [[0, 2], [2, 3]]], [3, true, []], [5, false, [[1, 4]]]])
    expect(rows.take()).toEqual([])
  })

  it("a row not above the previous one is malformed, at its position", () => {
    const rows = new RowBuilder(limits(), rules())
    expect(rows.openRow(4, false, 10)).toBeUndefined()
    expect(rows.openRow(4, false, 20)).toEqual(malformed("row 5 after row 5", 20))
  })

  it("a row index at maxRows, a column at maxColumns and a cell past maxCellsPerSheet fail", () => {
    const rows = new RowBuilder(limits({ maxRows: 10, maxColumns: 4, maxCellsPerSheet: 2 }), rules())
    expect(rows.openRow(10, false, 0)).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 10, actual: 11 })
    expect(rows.openRow(0, false, 0)).toBeUndefined()
    expect(rows.placeCell(4, 0)).toMatchObject({ _tag: "LimitExceeded", limit: "maxColumns", max: 4, actual: 5 })
    expect(add(rows, Cell.number(0, 1))).toBeUndefined()
    expect(add(rows, Cell.number(0, 2))).toBeUndefined()
    expect(add(rows, Cell.number(0, 3))).toMatchObject({ _tag: "LimitExceeded", limit: "maxCellsPerSheet", max: 2, actual: 3 })
  })

  // indexes go up, so the index check catches a row past maxRows first; only
  // a fractional index reaches the row-count guard
  it("a row past maxRows rows fails at its header, after the rows before it", () => {
    const rows = new RowBuilder(limits({ maxRows: 2 }), rules())
    expect(rows.openRow(0, false, 0)).toBeUndefined()
    expect(rows.openRow(1, false, 0)).toBeUndefined()
    expect(rows.openRow(1.5, false, 0)).toMatchObject({ limit: "maxRows", actual: 3 })
    expect(rows.take().map((r) => r.index)).toEqual([0, 1])
  })

  it("a cell outside a row is the format's error, or ignored", () => {
    const failing = new RowBuilder(limits(), rules())
    expect(failing.beginCell(7)).toEqual(malformed("cell outside a row", 7))
    const ignoring = new RowBuilder(limits(), rules({ cellOutsideRow: "ignore" }))
    expect(ignoring.beginCell(7)).toBe("skip")
    expect(ignoring.openRow(0, false, 0)).toBeUndefined()
    expect(ignoring.beginCell(7)).toBe("read")
    expect(ignoring.closeRow()).toBeUndefined()
    expect(ignoring.beginCell(7)).toBe("skip")
  })

  it("a cell without a column takes the previous cell's + 1, 0 at a row's start", () => {
    const rows = new RowBuilder(limits(), rules())
    expect(rows.openRow(0, false, 0)).toBeUndefined()
    expect(rows.nextColumn).toBe(0)
    expect(add(rows, Cell.number(5, 1))).toBeUndefined()
    expect(rows.nextColumn).toBe(6)
    expect(rows.openRow(1, false, 0)).toBeUndefined()
    expect(rows.nextColumn).toBe(0)
  })

  it("a row without an index takes its first cell's row, else the previous row + 1", () => {
    const rows = new RowBuilder(limits(), rules())
    // numbered by its first cell's own row; later cells without a row follow
    expect(rows.openRow(undefined, true, 0)).toBeUndefined()
    expect(rows.beginCell(1)).toBe("read")
    expect(rows.placeCell(2, 1, 4)).toBeUndefined()
    expect(rows.addCell(Cell.number(2, 1))).toBeUndefined()
    expect(rows.placeCell(3, 2)).toBeUndefined()
    expect(rows.addCell(Cell.number(3, 2))).toBeUndefined()
    // numbered by its first cell, which names no row: the previous row + 1
    expect(rows.openRow(undefined, false, 3)).toBeUndefined()
    expect(rows.placeCell(0, 4)).toBeUndefined()
    expect(rows.addCell(Cell.number(0, 3))).toBeUndefined()
    // without cells: numbered when it closes
    expect(rows.openRow(undefined, false, 5)).toBeUndefined()
    expect(rows.closeRow()).toBeUndefined()
    // a row with an index after one without
    expect(rows.openRow(9, false, 6)).toBeUndefined()
    expect(rows.closeRow()).toBeUndefined()
    expect(shape(rows)).toEqual([[4, true, [[2, 1], [3, 2]]], [5, false, [[0, 3]]], [6, false, []], [9, false, []]])
  })

  it("a row numbered by its first cell must still go up, and stay below maxRows", () => {
    const rows = new RowBuilder(limits({ maxRows: 8 }), rules())
    expect(rows.openRow(3, false, 0)).toBeUndefined()
    expect(rows.openRow(undefined, false, 1)).toBeUndefined()
    expect(rows.placeCell(0, 2, 3)).toEqual(malformed("row 4 after row 4", 2))
    const high = new RowBuilder(limits({ maxRows: 8 }), rules())
    expect(high.openRow(undefined, false, 0)).toBeUndefined()
    expect(high.placeCell(0, 1, 8)).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 8, actual: 9 })
    // a row without an index or cells after the last row allowed
    const last = new RowBuilder(limits({ maxRows: 8 }), rules())
    expect(last.openRow(7, false, 0)).toBeUndefined()
    expect(last.openRow(undefined, false, 1)).toBeUndefined()
    expect(last.closeRow()).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 8, actual: 9 })
    expect(last.take().map((r) => r.index)).toEqual([7])
  })

  it("a cell naming another row than its open row is the format's error; one naming the same row is fine", () => {
    const rows = new RowBuilder(limits(), rules())
    expect(rows.openRow(2, false, 0)).toBeUndefined()
    expect(rows.placeCell(0, 1, 2)).toBeUndefined()
    expect(rows.addCell(Cell.number(0, 1))).toBeUndefined()
    expect(rows.placeCell(1, 7, 5)).toEqual(malformed("cell in row 3 names another row", 7))
    // in a row numbered by its first cell, a later cell must agree with it
    expect(rows.openRow(undefined, false, 0)).toBeUndefined()
    expect(rows.placeCell(0, 1, 4)).toBeUndefined()
    expect(rows.addCell(Cell.number(0, 1))).toBeUndefined()
    expect(rows.placeCell(1, 8, 6)).toEqual(malformed("cell in row 5 names another row", 8))
  })

  it("a row repeating its columns holds at most 2 x maxColumns cells, and keeps the last of each", () => {
    const rows = new RowBuilder(limits({ maxColumns: 4 }), rules())
    expect(rows.openRow(0, false, 0)).toBeUndefined()
    let most = 0
    // progress: k up to 100
    for (let k = 0; k < 100; k++) {
      expect(add(rows, Cell.number(3 - (k % 4), k))).toBeUndefined()
      most = Math.max(most, rows.held)
    }
    expect(most).toBeLessThanOrEqual(8)
    expect(most).toBeGreaterThan(4)
    expect(rows.closeRow()).toBeUndefined()
    expect(rows.held).toBe(0)
    expect(shape(rows)).toEqual([[0, false, [[0, 99], [1, 98], [2, 97], [3, 96]]]])
  })
})

describe("held rows", () => {
  const held = (rows: HeldRows) => rows.take().map((r) => [r.index, r.hidden, r.cells.map((c) => [c.column, Cell.value(c)])])

  it("holds each row's cells until it is finished, in any row order; sorted, the last of a column kept", () => {
    const rows = new HeldRows(limits())
    expect(rows.addCell(5, Cell.number(2, 1))).toBeUndefined()
    expect(rows.addCell(2, Cell.number(1, 2))).toBeUndefined()
    expect(rows.addCell(5, Cell.number(0, 3))).toBeUndefined()
    expect(rows.addCell(5, Cell.number(2, 4))).toBeUndefined()
    expect(rows.held).toBe(4)
    expect(Array.from(rows.indexes())).toEqual([2, 5])
    expect(rows.ready).toBe(false)
    expect(rows.finishRow(2, true)).toBe(true)
    // a row without cells: nothing happens until the format sends it out without any
    expect(rows.finishRow(3, false)).toBe(false)
    expect(rows.held).toBe(3)
    expect(rows.ready).toBe(true)
    expect(rows.finishEmptyRow(3, false)).toBeUndefined()
    expect(rows.finishRow(5, false)).toBe(true)
    // a finished row holds no cells any more
    expect(rows.finishRow(5, false)).toBe(false)
    expect(rows.held).toBe(0)
    expect(held(rows)).toEqual([[2, true, [[1, 2]]], [3, false, []], [5, false, [[0, 3], [2, 4]]]])
    expect(rows.ready).toBe(false)
    expect(rows.take()).toEqual([])
  })

  it("a column at maxColumns and a cell past maxCellsPerSheet fail as they are added; a row at maxRows as it is finished", () => {
    const rows = new HeldRows(limits({ maxRows: 10, maxColumns: 4, maxCellsPerSheet: 2 }))
    expect(rows.addCell(0, Cell.number(4, 1))).toMatchObject({ _tag: "LimitExceeded", limit: "maxColumns", max: 4, actual: 5 })
    expect(rows.addCell(0, Cell.number(0, 1))).toBeUndefined()
    expect(rows.addCell(12, Cell.number(0, 2))).toBeUndefined()
    expect(rows.addCell(1, Cell.number(0, 3))).toMatchObject({ _tag: "LimitExceeded", limit: "maxCellsPerSheet", max: 2, actual: 3 })
    expect(rows.finishRow(0, false)).toBe(true)
    expect(rows.finishRow(12, false)).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 10, actual: 13 })
    expect(rows.finishEmptyRow(10, false)).toMatchObject({ _tag: "LimitExceeded", limit: "maxRows", max: 10, actual: 11 })
    expect(rows.take().map((r) => r.index)).toEqual([0])
  })

  it("a row repeating its columns holds at most 2 x maxColumns cells; the work counts cells, sorting and rows", () => {
    const rows = new HeldRows(limits({ maxColumns: 4 }))
    let most = 0
    // progress: k up to 100
    for (let k = 0; k < 100; k++) {
      expect(rows.addCell(k % 2, Cell.number(3 - (k % 4), k))).toBeUndefined()
      most = Math.max(most, rows.held)
    }
    // two rows, each compacted past 8 cells
    expect(most).toBeLessThanOrEqual(2 * 8)
    expect(rows.sorted).toBeGreaterThan(0)
    expect(rows.work).toBe(100 + rows.sorted)
    const sorted = rows.sorted
    expect(rows.finishRow(0, false)).toBe(true)
    expect(rows.finishRow(1, false)).toBe(true)
    expect(rows.held).toBe(0)
    expect(rows.work).toBe(100 + rows.sorted + 2)
    // both rows were left unsorted, and were sorted when finished
    expect(rows.sorted).toBeGreaterThan(sorted)
    expect(held(rows)).toEqual([[0, false, [[1, 98], [3, 96]]], [1, false, [[0, 99], [2, 97]]]])
  })
})

describe("layout builder", () => {
  it("hidden columns merge into sorted spans, clamped to maxColumns", () => {
    const layout = new LayoutBuilder(limits({ maxColumns: 10 }))
    layout.hideColumns(5, 6)
    layout.hideColumns(0, 1)
    layout.hideColumns(2, 2)
    layout.hideColumns(8, 100)
    layout.hideColumns(10, 12)
    expect(layout.result()).toEqual({ hiddenColumns: [{ first: 0, last: 2 }, { first: 5, last: 6 }, { first: 8, last: 9 }], merges: [] })
  })

  it("column records count against maxColumns, merges against maxCellsPerSheet", () => {
    const layout = new LayoutBuilder(limits({ maxColumns: 2, maxCellsPerSheet: 1 }))
    expect(layout.columnRecord()).toBeUndefined()
    expect(layout.columnRecord()).toBeUndefined()
    expect(layout.columnRecord()).toMatchObject({ _tag: "LimitExceeded", limit: "maxColumns", max: 2, actual: 3 })
    const range = { firstRow: 0, firstColumn: 0, lastRow: 1, lastColumn: 1 }
    expect(layout.merge(range)).toBeUndefined()
    expect(layout.merge(range)).toMatchObject({ _tag: "LimitExceeded", limit: "maxCellsPerSheet", max: 1, actual: 2 })
  })

  it("the last dimension wins; none is left out", () => {
    const layout = new LayoutBuilder(limits())
    expect("dimension" in layout.result()).toBe(false)
    layout.dimension({ firstRow: 0, firstColumn: 0, lastRow: 1, lastColumn: 1 })
    layout.dimension({ firstRow: 2, firstColumn: 2, lastRow: 3, lastColumn: 3 })
    expect(layout.result().dimension).toEqual({ firstRow: 2, firstColumn: 2, lastRow: 3, lastColumn: 3 })
  })
})

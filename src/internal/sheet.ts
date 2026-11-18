/**
 * Format-neutral helpers for the sheet readers: a row's cells in column
 * order, and hidden-column spans.
 *
 * @internal
 */
import type * as Cell from "../Cell.ts"
import type { ColumnSpan, Row } from "../Workbook.ts"

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

/** Sorts spans and merges overlapping or adjacent ones. */
export const mergeSpans = (spans: ReadonlyArray<ColumnSpan>): Array<ColumnSpan> => {
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

/** No rows: what a step that completed none emits. */
export const NO_ROWS: ReadonlyArray<Row> = []

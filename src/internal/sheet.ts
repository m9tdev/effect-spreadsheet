/**
 * Format-neutral helpers for the sheet readers: a row's cells in column
 * order, hidden-column spans, and the state of a pulled row stream.
 *
 * @internal
 */
import type * as Cell from "../Cell.ts"
import type { SpreadsheetError } from "../SpreadsheetError.ts"
import type { ColumnSpan } from "../Workbook.ts"

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

/**
 * Where a sheet's row stream is: reading input; failing, with the rows
 * completed before the error already emitted; or done (the input ended,
 * or the error was raised).
 */
export type PullState =
  | { readonly _tag: "Reading" }
  | { readonly _tag: "Failing"; readonly error: SpreadsheetError }
  | { readonly _tag: "Done" }

export const READING: PullState = { _tag: "Reading" }
export const DONE: PullState = { _tag: "Done" }

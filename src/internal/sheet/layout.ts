/**
 * The format-neutral layout collector: a sheet's hidden columns, merged
 * ranges and declared dimension (`SheetDetails`), with their limits. A
 * format's details reader decodes its records or elements and hands them
 * over; cost and memory go with the spans and merges given.
 *
 * Limits: column-description records count against `maxColumns` (a sheet
 * has at most that many columns to describe), hidden or not; merges
 * against `maxCellsPerSheet`.
 *
 * @internal
 */
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { ColumnSpan, Range, SheetDetails } from "../../Workbook.ts"
import { mergeSpans } from "../sheet.ts"

export class LayoutBuilder {
  private readonly hidden: Array<ColumnSpan> = []
  private readonly merges: Array<Range> = []
  private columnRecords = 0
  private declared: Range | undefined
  private readonly limits: Limits

  constructor(limits: Limits) {
    this.limits = limits
  }

  /** One column-description record, hidden or not: fails past `maxColumns` of them. */
  columnRecord(): SpreadsheetError | undefined {
    const max = this.limits.maxColumns
    if (++this.columnRecords > max) return new LimitExceeded({ limit: "maxColumns", max, actual: this.columnRecords })
    return undefined
  }

  /** Columns `first`..`last` (0-based) are hidden; the part at or past `maxColumns` is dropped. */
  hideColumns(first: number, last: number): void {
    const end = Math.min(last, this.limits.maxColumns - 1)
    if (first <= end) this.hidden.push({ first, last: end })
  }

  /** A merged range: fails past `maxCellsPerSheet` of them. */
  merge(range: Range): SpreadsheetError | undefined {
    const merges = this.merges
    const max = this.limits.maxCellsPerSheet
    if (merges.length >= max) return new LimitExceeded({ limit: "maxCellsPerSheet", max, actual: merges.length + 1 })
    merges.push(range)
    return undefined
  }

  /** The declared dimension; a later one replaces it. */
  dimension(range: Range): void {
    this.declared = range
  }

  /** The layout: hidden columns as sorted, merged spans; merges in the order given. */
  result(): SheetDetails {
    const dimension = this.declared
    return {
      hiddenColumns: mergeSpans(this.hidden),
      merges: this.merges,
      ...(dimension === undefined ? {} : { dimension })
    }
  }
}

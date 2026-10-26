/**
 * One pass over a worksheet part for its layout: hidden columns
 * (`<cols><col min max hidden>`), merged ranges (`<mergeCells>`) and the
 * declared dimension. Cell data is tokenized but not collected, so the cost
 * is linear in the part's bytes and the memory in the spans and merges.
 *
 * Limits: column spans count against `maxColumns` (a sheet has at most
 * that many columns to describe), merges against `maxCellsPerSheet`.
 *
 * @internal
 */
import { Effect, type Stream } from "effect"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { ColumnSpan, Range, SheetDetails } from "../../Workbook.ts"
import { parse } from "../xml/parse.ts"
import type { XmlHandler } from "../xml/tokenizer.ts"
import { attr, isTrue, malformed } from "./parts.ts"
import { parseRange } from "./refs.ts"

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

export const sheetDetails = (
  bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
  part: string,
  limits: Limits
): Effect.Effect<SheetDetails, SpreadsheetError> =>
  Effect.suspend(() => {
    const hidden: Array<ColumnSpan> = []
    const merges: Array<Range> = []
    let colElements = 0
    let dimension: Range | undefined
    let inSheetData = false
    const handler: XmlHandler = {
      open: (name, attributes) => {
        if (name === "sheetData") inSheetData = true
        if (inSheetData) return undefined
        if (name === "dimension") {
          const ref = attr(attributes, "ref")
          const range = ref === undefined ? undefined : parseRange(ref)
          if (range !== undefined) dimension = range
        } else if (name === "col") {
          if (++colElements > limits.maxColumns) {
            return new LimitExceeded({ limit: "maxColumns", max: limits.maxColumns, actual: colElements })
          }
          if (!isTrue(attr(attributes, "hidden"))) return undefined
          const min = attr(attributes, "min")
          const max = attr(attributes, "max")
          if (min === undefined || max === undefined || !/^\d{1,9}$/.test(min) || !/^\d{1,9}$/.test(max)) {
            return malformed("bad <col> range", part)
          }
          const first = Number(min) - 1
          const last = Math.min(Number(max), limits.maxColumns) - 1
          if (first < 0 || last < first) return first < 0 ? malformed("bad <col> range", part) : undefined
          hidden.push({ first, last })
        } else if (name === "mergeCell") {
          const ref = attr(attributes, "ref")
          const range = ref === undefined ? undefined : parseRange(ref)
          if (range === undefined) return malformed(`bad merge range ${JSON.stringify((ref ?? "").slice(0, 40))}`, part)
          if (merges.length >= limits.maxCellsPerSheet) {
            return new LimitExceeded({ limit: "maxCellsPerSheet", max: limits.maxCellsPerSheet, actual: merges.length + 1 })
          }
          merges.push(range)
        }
        return undefined
      },
      close: (name) => {
        if (name === "sheetData") inSheetData = false
        return undefined
      },
      text: () => undefined
    }
    return Effect.map(parse(bytes, handler, limits, part), (): SheetDetails => ({
      hiddenColumns: mergeSpans(hidden),
      merges,
      ...(dimension === undefined ? {} : { dimension })
    }))
  })

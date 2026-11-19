/**
 * One pass over a worksheet part for its layout: hidden columns
 * (`<cols><col min max hidden>`), merged ranges (`<mergeCells>`) and the
 * declared dimension. Cell data is tokenized but not collected, so the cost
 * is linear in the part's bytes and the memory in the spans and merges.
 *
 * The part is read through the rows reader's rules table (`sheetRules`),
 * so the layout elements are those outside the sheet data, and a nesting
 * the rows reader rejects (a `<row>` inside a row, a `<sheetData>` inside
 * the sheet data) is MalformedFile here too.
 *
 * The layout is collected by the shared `LayoutBuilder`, which holds its
 * limits: `<col>` elements count against `maxColumns` (a sheet has at most
 * that many columns to describe), merges against `maxCellsPerSheet`.
 *
 * @internal
 */
import { Effect, type Stream } from "effect"
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { SheetDetails } from "../../Workbook.ts"
import { parseRange } from "../refs.ts"
import { LayoutBuilder } from "../sheet/layout.ts"
import { parse } from "../xml/parse.ts"
import * as Attr from "../xml/attributes.ts"
import { rulesHandler } from "../xml/rules.ts"
import { malformed, NINE_DIGITS } from "./parts.ts"
import { sheetRules } from "./sheetRules.ts"

export const sheetDetails = (
  bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
  part: string,
  limits: Limits
): Effect.Effect<SheetDetails, SpreadsheetError> =>
  Effect.suspend(() => {
    const layout = new LayoutBuilder(limits)
    // the rows reader's table, so a part the rows fail on as malformed fails here too
    const handler = rulesHandler(sheetRules(part, {
      // cell data isn't layout
      cell: { open: () => "skip" },
      dimension: {
        open: (attributes) => {
          const ref = Attr.string(attributes, "ref")
          const range = ref === undefined ? undefined : parseRange(ref)
          if (range !== undefined) layout.dimension(range)
        }
      },
      col: {
        open: (attributes) => {
          const counted = layout.columnRecord()
          if (counted !== undefined) return counted
          if (!Attr.bool(attributes, "hidden")) return undefined
          // a hidden span needs both ends: a missing or bad one is MalformedFile
          const min = Attr.uint(attributes, "min", NINE_DIGITS)
          const max = Attr.uint(attributes, "max", NINE_DIGITS)
          if (typeof min !== "number" || typeof max !== "number" || min === 0) return malformed("bad <col> range", part)
          // a reversed span hides nothing; the part at or past maxColumns is dropped
          layout.hideColumns(min - 1, max - 1)
          return undefined
        }
      },
      mergeCell: {
        open: (attributes) => {
          const ref = Attr.string(attributes, "ref")
          const range = ref === undefined ? undefined : parseRange(ref)
          if (range === undefined) return malformed(`bad merge range ${JSON.stringify((ref ?? "").slice(0, 40))}`, part)
          return layout.merge(range)
        }
      }
    }))
    return Effect.map(parse(bytes, handler, limits, part), () => layout.result())
  })

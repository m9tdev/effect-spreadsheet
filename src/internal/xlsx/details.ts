/**
 * One pass over a worksheet part for its layout: hidden columns
 * (`<cols><col min max hidden>`), merged ranges (`<mergeCells>`) and the
 * declared dimension. Cell data is tokenized but not collected, so the cost
 * is linear in the part's bytes and the memory in the spans and merges.
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
import type { XmlHandler } from "../xml/tokenizer.ts"
import { malformed, NINE_DIGITS } from "./parts.ts"

export const sheetDetails = (
  bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
  part: string,
  limits: Limits
): Effect.Effect<SheetDetails, SpreadsheetError> =>
  Effect.suspend(() => {
    const layout = new LayoutBuilder(limits)
    let inSheetData = false
    let inRow = false
    const handler: XmlHandler = {
      open: (name, attributes, element) => {
        if (name === "sheetData") {
          // a nested one would end the sheet data at its end tag, reading what follows as layout
          if (inSheetData) return malformed("<sheetData> inside a <sheetData>", part)
          inSheetData = true
          return undefined
        }
        if (inSheetData) {
          if (name === "row") inRow = true
          // as in the rows reader, a cell's content is skipped whatever its
          // names, and only the cell's own end tag closes it
          else if (name === "c" && inRow) element.skip()
          return undefined
        }
        if (name === "dimension") {
          const ref = Attr.string(attributes, "ref")
          const range = ref === undefined ? undefined : parseRange(ref)
          if (range !== undefined) layout.dimension(range)
        } else if (name === "col") {
          const counted = layout.columnRecord()
          if (counted !== undefined) return counted
          if (!Attr.bool(attributes, "hidden")) return undefined
          // a hidden span needs both ends: a missing or bad one is MalformedFile
          const min = Attr.uint(attributes, "min", NINE_DIGITS)
          const max = Attr.uint(attributes, "max", NINE_DIGITS)
          if (typeof min !== "number" || typeof max !== "number" || min === 0) return malformed("bad <col> range", part)
          // a reversed span hides nothing; the part at or past maxColumns is dropped
          layout.hideColumns(min - 1, max - 1)
        } else if (name === "mergeCell") {
          const ref = Attr.string(attributes, "ref")
          const range = ref === undefined ? undefined : parseRange(ref)
          if (range === undefined) return malformed(`bad merge range ${JSON.stringify((ref ?? "").slice(0, 40))}`, part)
          return layout.merge(range)
        }
        return undefined
      },
      close: (name) => {
        if (name === "row") inRow = false
        else if (name === "sheetData") inSheetData = false
        return undefined
      }
    }
    return Effect.map(parse(bytes, handler, limits, part), () => layout.result())
  })

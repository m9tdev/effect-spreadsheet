/**
 * Builds a workbook from one source per sheet: a format reader says, for
 * each sheet, how to read its cells, that it holds none, or why it can't be
 * read, and gets the sheet list, the "no cells" and "broken" cases and the
 * sheet-to-source join from here.
 *
 * @internal
 */
import { Effect, Stream } from "effect"
import { SheetNotFound, type SpreadsheetError } from "../SpreadsheetError.ts"
import type { Limits } from "../SpreadsheetLimits.ts"
import { type DefinedName, type Format, make, type Row, type SheetDetails, type SheetInfo, type Visibility, type Workbook } from "../Workbook.ts"
import { NO_DETAILS } from "./sheet/layout.ts"

/**
 * How a sheet's rows and details are read. Each function gets the
 * effective limits of one read and must use exactly those.
 * - `Cells`: a sheet with cells.
 * - `NoCells`: a sheet that holds no cells (a chart or macro sheet): no
 *   rows, no details.
 * - `Broken`: a sheet that can't be read; reading its rows or details fails
 *   with `error`, while the workbook and its other sheets still open.
 */
export type SheetSource =
  | {
    readonly _tag: "Cells"
    readonly rows: (limits: Limits) => Stream.Stream<Row, SpreadsheetError>
    readonly details: (limits: Limits) => Effect.Effect<SheetDetails, SpreadsheetError>
  }
  | { readonly _tag: "NoCells" }
  | { readonly _tag: "Broken"; readonly error: SpreadsheetError }

/** A sheet as a reader gives it: its name, its visibility and how to read it. */
export interface SourcedSheet {
  readonly name: string
  readonly visibility: Visibility
  readonly source: SheetSource
}

/**
 * A workbook whose sheets are `sheets`, in order: each `SheetInfo.index` is
 * the sheet's position.
 */
export const fromSources = (options: {
  readonly format: Format
  readonly sheets: ReadonlyArray<SourcedSheet>
  readonly definedNames?: ReadonlyArray<DefinedName> | undefined
  readonly limits: Limits
}): Workbook => {
  const sources = options.sheets.map((s) => s.source)
  const sheets: Array<SheetInfo> = options.sheets.map((s, index) => ({ name: s.name, index, visibility: s.visibility }))
  // Each listed sheet's index is its position in `sources`. The workbook only
  // reads the sheets it lists (`Workbook.findSheet` returns one of `sheets`),
  // so the fallback, for a sheet of another workbook, is never reached.
  const sourceOf = (sheet: SheetInfo): SheetSource =>
    sources[sheet.index] ?? { _tag: "Broken", error: new SheetNotFound({ sheet: sheet.name }) }
  return make({
    format: options.format,
    sheets,
    definedNames: options.definedNames,
    limits: options.limits,
    rows: (sheet, limits) => {
      const source = sourceOf(sheet)
      switch (source._tag) {
        case "Cells":
          return source.rows(limits)
        case "NoCells":
          return Stream.empty
        case "Broken":
          return Stream.fail(source.error)
      }
    },
    details: (sheet, limits) => {
      const source = sourceOf(sheet)
      switch (source._tag) {
        case "Cells":
          return source.details(limits)
        case "NoCells":
          return Effect.succeed(NO_DETAILS)
        case "Broken":
          return Effect.fail(source.error)
      }
    }
  })
}

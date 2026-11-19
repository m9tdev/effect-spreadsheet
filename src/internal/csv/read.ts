/**
 * The CSV reader behind `Csv.read` and `Spreadsheet.read`, for limits
 * already resolved and, for bytes, a size already checked.
 */
import { Effect, Stream } from "effect"
import * as Result from "effect/Result"
import type { ReadOptions } from "../../csv/index.ts"
import type { InvalidOption, SpreadsheetError } from "../../SpreadsheetError.ts"
import * as Workbook from "../../Workbook.ts"
import { stored, type ValidLimits } from "../input.ts"
import { csvOptions } from "./options.ts"
import { rowsFromByteStream } from "./reader.ts"

export const toStream = <E, R>(input: Uint8Array | Stream.Stream<Uint8Array, E, R>): Stream.Stream<Uint8Array, E, R> =>
  input instanceof Uint8Array ? Stream.succeed(input) : input

/**
 * A one-sheet workbook over `input`, after checking the options (the
 * delimiter, then the encoding). Parsing happens when rows are streamed,
 * with the options as they were checked here. `options.limits` is not
 * read: `limits` are the ones in effect.
 */
export const readCsvWorkbook = (
  input: Uint8Array | Stream.Stream<Uint8Array, SpreadsheetError>,
  limits: ValidLimits,
  options: ReadOptions | undefined
): Effect.Effect<Workbook.Workbook, InvalidOption> => {
  const checked = csvOptions(options)
  if (Result.isFailure(checked)) return Effect.fail(checked.failure)
  const csv = checked.success
  const sheet: Workbook.SheetInfo = { name: options?.sheetName ?? "Sheet1", index: 0, visibility: "visible" }
  return Effect.succeed(Workbook.make({
    format: "csv",
    sheets: [sheet],
    limits: stored(limits),
    rows: (_, effective) => rowsFromByteStream(toStream(input), csv, effective)
  }))
}

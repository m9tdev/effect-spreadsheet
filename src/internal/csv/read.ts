/**
 * The CSV reader behind `Csv.read` and `Spreadsheet.read`, for limits
 * already resolved and, for bytes, a size already checked.
 */
import { Effect, Stream } from "effect"
import type { ReadOptions } from "../../csv/index.ts"
import type { InvalidOption, SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import * as Workbook from "../../Workbook.ts"
import { stored, type ValidLimits } from "../input.ts"
import { validateEncoding } from "./decode.ts"
import { rowsFromByteStream, validateDelimiter } from "./reader.ts"

export const resolved = (options: ReadOptions | undefined, limits: Limits) => ({
  encoding: options?.encoding ?? "auto",
  delimiter: options?.delimiter,
  inferTypes: options?.inferTypes ?? false,
  strict: options?.strict ?? false,
  limits
})

export const toStream = <E, R>(input: Uint8Array | Stream.Stream<Uint8Array, E, R>): Stream.Stream<Uint8Array, E, R> =>
  input instanceof Uint8Array ? Stream.succeed(input) as Stream.Stream<Uint8Array, E, R> : input

/**
 * A one-sheet workbook over `input`, after checking the delimiter and
 * encoding options. Parsing happens when rows are streamed. `options.limits`
 * is not read: `limits` are the ones in effect.
 */
export const readCsvWorkbook = (
  input: Uint8Array | Stream.Stream<Uint8Array, SpreadsheetError>,
  limits: ValidLimits,
  options: ReadOptions | undefined
): Effect.Effect<Workbook.Workbook, InvalidOption> => {
  const invalid = validateDelimiter(options?.delimiter) ?? validateEncoding(options?.encoding ?? "auto")
  if (invalid !== undefined) return Effect.fail(invalid)
  const sheet: Workbook.SheetInfo = { name: options?.sheetName ?? "Sheet1", index: 0, visibility: "visible" }
  return Effect.succeed(Workbook.make({
    format: "csv",
    sheets: [sheet],
    limits: stored(limits),
    rows: (_, effective) => rowsFromByteStream(toStream(input), resolved(options, effective))
  }))
}

/**
 * CSV / TSV / delimited-text reader.
 *
 * ```ts
 * import * as Csv from "effect-spreadsheet/csv"
 *
 * Csv.rows(bytes, { inferTypes: true })   // Stream<Row, SpreadsheetError>
 * Csv.read(bytes)                         // Effect<Workbook, SpreadsheetError>, one sheet
 * ```
 *
 * Behaviour:
 * - Encoding: a BOM wins (UTF-8, UTF-16LE/BE). Without one, `"auto"` looks
 *   at the first 64 KiB: UTF-16 if every other byte is NUL, else UTF-8 if
 *   they are valid UTF-8, else windows-1252. Any `TextDecoder` label can
 *   be forced with `encoding`.
 * - Delimiter: `delimiter`, else an Excel `sep=X` first line (consumed),
 *   else sniffed from the first 16 Ki characters among `,` `\t` `;` `|`.
 * - Quoting is RFC 4180 (`""` escapes, delimiters and line breaks inside
 *   quotes). Lenient by default: text after a closing quote is appended,
 *   an unterminated quote runs to end of input. `strict: true` turns both
 *   into `MalformedFile`.
 * - Line breaks LF, CRLF and CR. Rows may have any number of fields.
 *   Empty fields produce no cell. Blank lines are rows without cells.
 * - Values are text unless `inferTypes: true` (see {@link ReadOptions}).
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import { validateEncoding } from "../internal/csv/decode.ts"
import { rowsFromByteStream, validateDelimiter } from "../internal/csv/reader.ts"
import { LimitExceeded, type SpreadsheetError } from "../SpreadsheetError.ts"
import * as SpreadsheetLimits from "../SpreadsheetLimits.ts"
import * as Workbook from "../Workbook.ts"

/**
 * `"auto"` (default) or a `TextDecoder` label.
 *
 * @since 0.1.0
 * @category models
 */
export type Encoding = "auto" | "utf-8" | "utf-16le" | "utf-16be" | "windows-1252" | (string & {})

/**
 * @since 0.1.0
 * @category models
 */
export interface ReadOptions {
  /** Default `"auto"`. */
  readonly encoding?: Encoding | undefined
  /** One character. Default: `sep=` line, else sniffed. */
  readonly delimiter?: string | undefined
  /**
   * Default `false`: every value is a `Text` cell. With `true`:
   * `TRUE`/`FALSE` (any case) become `Boolean`; plain decimals (optional
   * sign, fraction, exponent; no leading zeros, at most 15 significant
   * digits, no thousands separators, currency or `%`) become `Number`; the
   * Excel error literals (`#N/A`, `#DIV/0!`, ...) become `Error`. Dates and
   * formulas are never inferred (`=SUM(A1)` stays text).
   */
  readonly inferTypes?: boolean | undefined
  /** Default `false`. See the module docs. */
  readonly strict?: boolean | undefined
  /** Name of the single sheet {@link read} reports. Default `"Sheet1"`. */
  readonly sheetName?: string | undefined
  /** Overrides on top of the `SpreadsheetLimits` in context. */
  readonly limits?: SpreadsheetLimits.LimitOverrides | undefined
}

const resolved = (options: ReadOptions | undefined, limits: SpreadsheetLimits.Limits) => ({
  encoding: options?.encoding ?? "auto",
  delimiter: options?.delimiter,
  inferTypes: options?.inferTypes ?? false,
  strict: options?.strict ?? false,
  limits
})

const toStream = <E, R>(input: Uint8Array | Stream.Stream<Uint8Array, E, R>): Stream.Stream<Uint8Array, E, R> =>
  input instanceof Uint8Array ? Stream.succeed(input) as Stream.Stream<Uint8Array, E, R> : input

/**
 * Streams the rows of delimited text. A `Uint8Array` is read in 64 KiB
 * pieces; a byte stream may be chunked anyhow (chunk boundaries may split
 * characters, quotes or CRLF). The stream is re-runnable when the input is.
 *
 * @since 0.1.0
 * @category readers
 */
export const rows: {
  (input: Uint8Array, options?: ReadOptions): Stream.Stream<Workbook.Row, SpreadsheetError>
  <E, R>(input: Stream.Stream<Uint8Array, E, R>, options?: ReadOptions): Stream.Stream<Workbook.Row, E | SpreadsheetError, R>
} = (<E, R>(
  input: Uint8Array | Stream.Stream<Uint8Array, E, R>,
  options?: ReadOptions
): Stream.Stream<Workbook.Row, E | SpreadsheetError, R> =>
  Stream.unwrap(
    Effect.map(SpreadsheetLimits.resolve(options?.limits), (limits) =>
      rowsFromByteStream(toStream(input), resolved(options, limits)))
  )) as never

/**
 * Opens delimited text as a one-sheet workbook. The limits are resolved
 * here (context plus `options.limits`) and kept in the workbook; for a
 * `Uint8Array` the size limit is checked here too. Parsing happens when
 * rows are streamed.
 * A byte-stream input is consumed each time the rows are streamed, so it
 * must be re-runnable to stream them more than once, and may only fail
 * with `SpreadsheetError` (map other errors first, or use {@link rows},
 * which keeps the input's error type).
 *
 * @since 0.1.0
 * @category readers
 */
export const read = (
  input: Uint8Array | Stream.Stream<Uint8Array, SpreadsheetError>,
  options?: ReadOptions
): Effect.Effect<Workbook.Workbook, SpreadsheetError> =>
  Effect.gen(function*() {
    const limits = yield* SpreadsheetLimits.resolve(options?.limits)
    if (input instanceof Uint8Array && input.length > limits.maxInputBytes) {
      return yield* new LimitExceeded({ limit: "maxInputBytes", max: limits.maxInputBytes, actual: input.length })
    }
    const invalid = validateDelimiter(options?.delimiter) ?? validateEncoding(options?.encoding ?? "auto")
    if (invalid !== undefined) return yield* invalid
    const sheet: Workbook.SheetInfo = { name: options?.sheetName ?? "Sheet1", index: 0, visibility: "visible" }
    return Workbook.make({
      format: "csv",
      sheets: [sheet],
      limits,
      rows: (_, effective) => rowsFromByteStream(toStream(input), resolved(options, effective))
    })
  })

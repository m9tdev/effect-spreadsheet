/**
 * Entry points: open any supported file from its bytes, and turn rows into
 * records.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import type * as Cell from "./Cell.ts"
import * as Csv from "./csv/index.ts"
import { type RecordsOptions, recordsStream } from "./internal/records.ts"
import { detect } from "./internal/sniff.ts"
import { LimitExceeded, type SpreadsheetError, UnsupportedFormat } from "./SpreadsheetError.ts"
import * as SpreadsheetLimits from "./SpreadsheetLimits.ts"
import type { Format, Row, Workbook } from "./Workbook.ts"

export type {
  /** @since 0.1.0 */
  Format,
  /** @since 0.1.0 */
  RecordsOptions
}

/**
 * @since 0.1.0
 * @category models
 */
export interface ReadOptions {
  /** Overrides on top of the `SpreadsheetLimits` in context. */
  readonly limits?: SpreadsheetLimits.LimitOverrides | undefined
  /** Options used when the bytes turn out to be delimited text. */
  readonly csv?: Omit<Csv.ReadOptions, "limits"> | undefined
}

/**
 * Detects the format from the bytes alone (magic numbers and structure;
 * there is no file-name input, on purpose). Only zip containers are looked
 * into (their central directory), so this is cheap.
 *
 * @since 0.1.0
 * @category detection
 */
export const detectFormat = (
  bytes: Uint8Array,
  options?: { readonly limits?: SpreadsheetLimits.LimitOverrides | undefined }
): Effect.Effect<Format, SpreadsheetError> =>
  Effect.flatMap(SpreadsheetLimits.resolve(options?.limits), (limits) => {
    if (bytes.length > limits.maxInputBytes) {
      return Effect.fail(new LimitExceeded({ limit: "maxInputBytes", max: limits.maxInputBytes, actual: bytes.length }))
    }
    const format = detect(bytes, limits)
    return typeof format === "string" ? Effect.succeed(format) : Effect.fail(format)
  })

const unsupported: Partial<Record<Format, string>> = {
  xlsx: "XLSX/XLSM reading is not implemented yet",
  xlsb: "XLSB reading is not implemented yet",
  xls: "XLS (BIFF) reading is not implemented yet",
  cfb: "OLE compound file (legacy .xls, or an encrypted/password-protected OOXML workbook); not readable yet",
  ods: "ODS reading is not implemented",
  spreadsheetml: "SpreadsheetML 2003 reading is not implemented yet",
  sylk: "SYLK reading is not implemented",
  zip: "zip archive without a known workbook",
  html: "HTML is not a supported spreadsheet format",
  xml: "XML document that is not SpreadsheetML",
  pdf: "PDF is not a spreadsheet",
  image: "image file is not a spreadsheet",
  rtf: "RTF is not a spreadsheet",
  unknown: "binary data of unknown format"
}

/**
 * Opens a workbook from its bytes, whatever its format: the format is
 * detected from the content ({@link detectFormat}), never from a name.
 * Formats without a reader fail with `UnsupportedFormat` naming what was
 * found.
 *
 * @since 0.1.0
 * @category readers
 */
export const read = (bytes: Uint8Array, options?: ReadOptions): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.flatMap(detectFormat(bytes, options), (format): Effect.Effect<Workbook, SpreadsheetError> => {
    switch (format) {
      case "csv":
        return Csv.read(bytes, { ...options?.csv, limits: options?.limits })
      default:
        return Effect.fail(new UnsupportedFormat({ format, reason: unsupported[format] ?? "not supported" }))
    }
  })

/**
 * Opens delimited text (CSV/TSV/...). Same as `effect-spreadsheet/csv`'s
 * `read`.
 *
 * @since 0.1.0
 * @category readers
 */
export const readCsv: (
  input: Uint8Array | Stream.Stream<Uint8Array, SpreadsheetError>,
  options?: Csv.ReadOptions
) => Effect.Effect<Workbook, SpreadsheetError> = Csv.read

/**
 * Turns rows into records keyed by a header (one object per data row, raw
 * cell values):
 *
 * - `header: "first-row"`: the first row gives the names.
 * - `header: n`: the row with index `n` gives the names; rows before it are
 *   skipped.
 * - `header: [...]`: the given names, by column (repeats made unique the
 *   same way); every row is data and columns past the array are dropped.
 *
 * Header names are made unique (`name`, `name_1`, `name_2`, ...); an empty
 * or Empty header cell leaves its column unnamed.
 * A column with no header cell is named `__EMPTY_k` when a data cell first
 * appears in it, and `defval` fills only the header-named columns.
 *
 * Cost is bounded per run: every cell of the rows (Empty included) counts
 * against `maxCellsPerSheet`, every `defval` fill against `maxRecordFields`;
 * going over fails with `LimitExceeded` after the records built so far.
 * Records come out in batches of about 64 Ki keys, yielding in between, so a
 * wide header times many rows doesn't block the event loop. The limits come
 * from the `SpreadsheetLimits` in context where the records are produced
 * (plus `options.limits`), not from the workbook the rows came from.
 *
 * Values are `Cell.value`s (text, number, boolean, Date, error code).
 * Rows without values are skipped unless `blankRows: true`.
 *
 * On a stream it works lazily, keeping the stream's error and context; on
 * an iterable of rows it gives an Effect of the array.
 *
 * @since 0.1.0
 * @category records
 */
export const toRecords: {
  <E, R, D = never>(
    rows: Stream.Stream<Row, E, R>,
    options: RecordsOptions<D>
  ): Stream.Stream<Record<string, Cell.Value | D>, E | LimitExceeded, R>
  <D = never>(
    rows: Iterable<Row>,
    options: RecordsOptions<D>
  ): Effect.Effect<Array<Record<string, Cell.Value | D>>, LimitExceeded>
} = (<E, R>(rows: Stream.Stream<Row, E, R> | Iterable<Row>, options: RecordsOptions<unknown>) => {
  const stream = Stream.unwrap(
    Effect.map(SpreadsheetLimits.resolve(options.limits), (limits) =>
      recordsStream(
        Stream.isStream(rows) ? rows as Stream.Stream<Row, E, R> : Stream.fromIterable(rows as Iterable<Row>),
        options,
        limits
      ))
  )
  return Stream.isStream(rows) ? stream : Effect.map(Stream.runCollect(stream), (chunk) => Array.from(chunk))
}) as never

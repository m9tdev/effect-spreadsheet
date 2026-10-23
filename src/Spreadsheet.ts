/**
 * Entry points: open any supported file from its bytes, and turn rows into
 * records.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import * as Csv from "./csv/index.ts"
import { RecordBuilder, type RecordsOptions } from "./internal/records.ts"
import { detect, type Format } from "./internal/sniff.ts"
import { LimitExceeded, type SpreadsheetError, UnsupportedFormat } from "./SpreadsheetError.ts"
import * as SpreadsheetLimits from "./SpreadsheetLimits.ts"
import type { Row, Workbook } from "./Workbook.ts"

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
 * Turns rows into records keyed by a header, like SheetJS `sheet_to_json`
 * (object mode, raw values):
 *
 * - `header: "first-row"`: the first row gives the names.
 * - `header: n`: the row with index `n` gives the names; rows before it are
 *   skipped.
 * - `header: [...]`: the given names, by column; every row is data and
 *   columns past the array are dropped.
 *
 * Header names are made unique like SheetJS (`__EMPTY`, `name_1`, ...);
 * data columns past the header row are named `__EMPTY_k` as they appear.
 * Values are `Cell.value`s (text, number, boolean, Date, error code).
 * Rows without values are skipped unless `blankRows: true`.
 *
 * Works on a stream (lazily, keeping its error and context) or on an
 * array of rows.
 *
 * @since 0.1.0
 * @category records
 */
export const toRecords: {
  <E, R>(rows: Stream.Stream<Row, E, R>, options: RecordsOptions): Stream.Stream<Record<string, unknown>, E, R>
  (rows: Iterable<Row>, options: RecordsOptions): Array<Record<string, unknown>>
} = (<E, R>(rows: Stream.Stream<Row, E, R> | Iterable<Row>, options: RecordsOptions) => {
  if (Stream.isStream(rows)) {
    return Stream.suspend(() => {
      const builder = new RecordBuilder(options)
      return (rows as Stream.Stream<Row, E, R>).pipe(
        Stream.map((row) => builder.push(row)),
        Stream.filter((record): record is Record<string, unknown> => record !== undefined)
      )
    })
  }
  const builder = new RecordBuilder(options)
  const out: Array<Record<string, unknown>> = []
  for (const row of rows as Iterable<Row>) {
    const record = builder.push(row)
    if (record !== undefined) out.push(record)
  }
  return out
}) as never

/**
 * Entry points: open any supported file from its bytes, and turn rows into
 * records.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import type * as Cell from "./Cell.ts"
import * as Csv from "./csv/index.ts"
import { RecordBuilder, type RecordsOptions } from "./internal/records.ts"
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
 * Turns rows into records keyed by a header, like SheetJS `sheet_to_json`
 * (object mode, raw values):
 *
 * - `header: "first-row"`: the first row gives the names.
 * - `header: n`: the row with index `n` gives the names; rows before it are
 *   skipped.
 * - `header: [...]`: the given names, by column; every row is data and
 *   columns past the array are dropped.
 *
 * Header names are made unique like SheetJS (`__EMPTY`, `name_1`, ...).
 * A column with no header cell is named `__EMPTY_k` when a data cell first
 * appears in it, and `defval` fills only the header-named columns, so the
 * cost follows the cells present. Every key written (fills included) counts
 * against `maxCellsPerSheet` for the run; going over fails with
 * `LimitExceeded`.
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
  if (Stream.isStream(rows)) {
    return Stream.unwrap(Effect.map(SpreadsheetLimits.resolve(options.limits), (limits) =>
      Stream.suspend(() => {
        const builder = new RecordBuilder(options, limits.maxCellsPerSheet)
        return (rows as Stream.Stream<Row, E, R>).pipe(
          Stream.mapArray((chunk) => [pushAll(builder, chunk)] as const),
          Stream.flatMap((out) => Array.isArray(out) ? Stream.fromArray(out) : Stream.fail(out as LimitExceeded))
        )
      })))
  }
  return Effect.flatMap(SpreadsheetLimits.resolve(options.limits), (limits) => {
    const out = pushAll(new RecordBuilder(options, limits.maxCellsPerSheet), rows as Iterable<Row>)
    return Array.isArray(out) ? Effect.succeed(out) : Effect.fail(out as LimitExceeded)
  })
}) as never

const pushAll = (builder: RecordBuilder, rows: Iterable<Row>): Array<Record<string, unknown>> | LimitExceeded => {
  const out: Array<Record<string, unknown>> = []
  // progress: one row per iteration
  for (const row of rows) {
    const record = builder.push(row)
    if (record instanceof LimitExceeded) return record
    if (record !== undefined) out.push(record)
  }
  return out
}

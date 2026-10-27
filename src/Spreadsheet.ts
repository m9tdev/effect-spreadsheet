/**
 * Entry points: open any supported file from its bytes, and turn rows into
 * records.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import type * as Cell from "./Cell.ts"
import * as Csv from "./csv/index.ts"
import * as Xls from "./xls/index.ts"
import * as Xlsb from "./xlsb/index.ts"
import * as Xlsx from "./xlsx/index.ts"
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
 * there is no file-name input, on purpose). Containers are looked into, so
 * this stays cheap: a zip's central directory, and a compound file's
 * directory (`"xls"` if it has a Workbook or Book stream, `"cfb"` if not).
 * A broken container fails with a typed error (`MalformedFile`,
 * `LimitExceeded`) rather than giving a format.
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
  cfb: "OLE compound file without a workbook stream",
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
 * The workbook keeps using `bytes` after this returns: XLSX and XLSB zip entries
 * and XLS workbook streams (bare, or in consecutive sectors of a compound
 * file) are views of them, not copies, read when a sheet's rows are. Don't
 * change them while the workbook is in use; pass a copy if the buffer is
 * reused.
 *
 * @since 0.1.0
 * @category readers
 */
export const read = (bytes: Uint8Array, options?: ReadOptions): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.flatMap(detectFormat(bytes, options), (format): Effect.Effect<Workbook, SpreadsheetError> => {
    switch (format) {
      case "csv":
        return Csv.read(bytes, { ...options?.csv, limits: options?.limits })
      case "xlsx":
        return Xlsx.read(bytes, { limits: options?.limits })
      case "xlsb":
        return Xlsb.read(bytes, { limits: options?.limits })
      case "xls":
      case "cfb": // an encrypted OOXML package (Encrypted), or no workbook stream
        return Xls.read(bytes, { limits: options?.limits })
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
 *   same way); every row is data. Columns past the array, and columns whose
 *   entry is a hole, `undefined` or `null`, are dropped. An entry `""` is
 *   treated like a blank header cell (an `__EMPTY…` key when data appears,
 *   no `defval` fill).
 *
 * Header names are made unique (`name`, `name_1`, `name_2`, ...). A column
 * without a header name (a blank or Empty header cell, a gap in the header
 * row, a column past it) gets its key when a data cell first appears in it:
 * the first such key is `__EMPTY`, then `__EMPTY_1`, `__EMPTY_2`, and so on.
 * `defval` fills only the header-named columns.
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

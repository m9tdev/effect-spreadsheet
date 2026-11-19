/**
 * Entry points: open any supported file from its bytes, and turn rows into
 * records.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import type * as Cell from "./Cell.ts"
import * as Csv from "./csv/index.ts"
import { readCsvWorkbook } from "./internal/csv/read.ts"
import { type Accepted, accept } from "./internal/input.ts"
import { WorkMeter } from "./internal/meter.ts"
import { type Detected, detect, probe } from "./internal/probe.ts"
import { type RecordsOptions, recordsStream } from "./internal/records.ts"
import { readProbed } from "./internal/xls/read.ts"
import { readXlsbPackage } from "./internal/xlsb/read.ts"
import { readXlsxPackage } from "./internal/xlsx/read.ts"
import { type InvalidOption, LimitExceeded, type SpreadsheetError, UnsupportedFormat } from "./SpreadsheetError.ts"
import * as SpreadsheetLimits from "./SpreadsheetLimits.ts"
import type { Format, ReadableFormat, Row, UnreadableFormat, Workbook } from "./Workbook.ts"

export type {
  /** @since 0.1.0 */
  Format,
  /** @since 0.1.0 */
  ReadableFormat,
  /** @since 0.1.0 */
  RecordsOptions,
  /** @since 0.1.0 */
  UnreadableFormat
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

/** The format of an accepted input. */
const detected = (input: Accepted): Effect.Effect<Format, SpreadsheetError> => detect(input.bytes, input.limits)

/**
 * Detects the format from the bytes alone (magic numbers and structure;
 * there is no file-name input, on purpose). Containers are looked into as
 * far as reading them would: a compound file's directory (`"xls"` if it has
 * a Workbook or Book stream, `"cfb"` if not), and a zip's central
 * directory. A zip that looks like an OOXML package (it has `_rels/.rels`,
 * `xl/workbook.xml` or `xl/workbook.bin`) is opened, and its workbook part
 * found through the package relationships, as the XLSX and XLSB readers
 * find it: `"xlsx"` or `"xlsb"` by that part, `"zip"` if it isn't a
 * workbook. A broken container fails with a typed error (`MalformedFile`,
 * `LimitExceeded`, or what reading the package relationships fails with)
 * rather than giving a format.
 *
 * @since 0.1.0
 * @category detection
 */
export const detectFormat = (
  bytes: Uint8Array,
  options?: { readonly limits?: SpreadsheetLimits.LimitOverrides | undefined }
): Effect.Effect<Format, SpreadsheetError> =>
  accept(bytes, options?.limits, detected)

/**
 * Why each format detection hands on without a reader fails: every
 * `Detected` format but `"csv"`, the one read by format (compound files,
 * bare BIFF streams and packages go by what the probe found in them). A
 * new detected format doesn't compile until it gets a reason here.
 */
const unsupported: Record<Exclude<Detected["format"], "csv">, string> = {
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
 * found. The options are checked here and kept: changing the options object
 * later doesn't affect the workbook's rows.
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
  // suspend: the options are read when the effect runs, each time it runs
  Effect.suspend(() =>
    accept(bytes, options?.limits, (input) =>
      Effect.flatMap(WorkMeter, (meter) =>
        Effect.flatMap(probe(input.bytes, input.limits, meter), (probed): Effect.Effect<Workbook, SpreadsheetError> => {
          switch (probed._tag) {
            case "Package":
              return probed.kind === "xlsx"
                ? readXlsxPackage(probed.archive, probed.workbook, input.limits)
                : readXlsbPackage(probed.archive, probed.workbook, input.limits)
            case "Detected":
              return readDetected(probed.format, input, options)
            default:
              return readProbed(probed, input.limits)
          }
        })))
  )

/** Reads an input by the format detection gave it; formats without a reader fail. */
const readDetected = (
  format: Detected["format"],
  input: Accepted,
  options: ReadOptions | undefined
): Effect.Effect<Workbook, SpreadsheetError> => {
  return format === "csv"
    ? readCsvWorkbook(input.bytes, input.limits, options?.csv)
    : Effect.fail(new UnsupportedFormat({ format, reason: unsupported[format] }))
}

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
 * - `header: n`: the row with index `n` (a safe integer, 0 or more) gives the
 *   names; rows before it are skipped.
 * - `header: [...]`: the given names, by column (repeats made unique the
 *   same way); every row is data. Columns past the array, and columns whose
 *   entry is a hole, `undefined` or `null`, are dropped. An entry `""` is
 *   treated like a blank header cell (an `__EMPTY…` key when data appears,
 *   no `defval` fill).
 *
 * Any other `header` (another string, a fraction, a negative number, `NaN`,
 * an array entry that is not a string, `null` or `undefined`) fails with
 * `InvalidOption` before any row is read.
 *
 * Header names are made unique (`name`, `name_1`, `name_2`, ...). A column
 * without a header name (a blank or Empty header cell, a gap in the header
 * row, a column past it) gets its key when a data cell first appears in it
 * (an Empty cell, such as a formatted blank cell, is not data): the first
 * such key is `__EMPTY`, then `__EMPTY_1`, `__EMPTY_2`, and so on. `defval`
 * fills only the header-named columns.
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
 * Rows without values are skipped unless `blankRows: true`. A row with more
 * than one cell in a column (the readers never make one) takes the last,
 * also when it is Empty, like the cells the readers keep.
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
  ): Stream.Stream<Record<string, Cell.Value | D>, E | LimitExceeded | InvalidOption, R>
  <D = never>(
    rows: Iterable<Row>,
    options: RecordsOptions<D>
  ): Effect.Effect<Array<Record<string, Cell.Value | D>>, LimitExceeded | InvalidOption>
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

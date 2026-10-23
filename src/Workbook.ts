/**
 * The format-neutral workbook model: sheets, rows, and a way to stream a
 * sheet's rows.
 *
 * @since 0.1.0
 */
import { Effect, Schema, Stream } from "effect"
import { Cell } from "./Cell.ts"
import { SheetNotFound, type SpreadsheetError } from "./SpreadsheetError.ts"
import { type LimitOverrides, type Limits, merge } from "./SpreadsheetLimits.ts"

/**
 * One row. `index` is the 0-based row index in the sheet (rows that hold
 * nothing may be skipped by a format, so indexes can jump). `cells` holds
 * only the cells present, in column order, each with its own `column`.
 * `hidden` reports a hidden row; hidden rows are always emitted.
 *
 * @since 0.1.0
 * @category schemas
 */
export const Row = Schema.Struct({
  index: Schema.Int,
  hidden: Schema.Boolean,
  cells: Schema.Array(Cell)
})

/**
 * @since 0.1.0
 * @category models
 */
export interface Row {
  readonly index: number
  readonly hidden: boolean
  readonly cells: ReadonlyArray<Cell>
}

/**
 * A cell range, 0-based and inclusive.
 *
 * @since 0.1.0
 * @category schemas
 */
export const Range = Schema.Struct({
  firstRow: Schema.Int,
  firstColumn: Schema.Int,
  lastRow: Schema.Int,
  lastColumn: Schema.Int
})

/** @since 0.1.0 @category models */
export type Range = typeof Range.Type

/**
 * @since 0.1.0
 * @category schemas
 */
export const Visibility = Schema.Literals(["visible", "hidden", "veryHidden"])

/** @since 0.1.0 @category models */
export type Visibility = typeof Visibility.Type

/**
 * A sheet's metadata. `dimension` is the range the file declares, when it
 * declares one; it is informative only (readers never size anything by it).
 *
 * @since 0.1.0
 * @category schemas
 */
export const SheetInfo = Schema.Struct({
  name: Schema.String,
  index: Schema.Int,
  visibility: Visibility,
  dimension: Schema.optionalKey(Range)
})

/** @since 0.1.0 @category models */
export type SheetInfo = typeof SheetInfo.Type

/**
 * A file format, as detected from the bytes (`Spreadsheet.detectFormat`)
 * and as reported by `Workbook.format`. Spreadsheet formats are named after
 * the reader that handles them; the others say what was found instead.
 *
 * @since 0.1.0
 * @category models
 */
export type Format =
  | "csv"
  | "xlsx"
  | "xlsb"
  | "xls"
  | "ods"
  | "spreadsheetml"
  | "sylk"
  /** OLE compound file: legacy .xls, or an encrypted OOXML workbook (told apart once the CFB reader exists). */
  | "cfb"
  /** A zip that isn't a known workbook. */
  | "zip"
  | "html"
  | "xml"
  | "pdf"
  | "image"
  | "rtf"
  /** Binary data of no known kind. */
  | "unknown"

/**
 * Options for {@link rows}.
 *
 * @since 0.1.0
 * @category models
 */
export interface RowsOptions {
  /**
   * Overrides on top of the limits the workbook was read with (the
   * `SpreadsheetLimits` in context at read time, plus the read call's own
   * overrides). The context at streaming time doesn't matter.
   */
  readonly limits?: LimitOverrides | undefined
}

/** @internal */
export const RowsSource: unique symbol = Symbol.for("effect-spreadsheet/Workbook/rows") as never

/**
 * A workbook opened by a reader. Its sheets are known up front; rows are
 * read lazily, sheet by sheet, through {@link rows}.
 *
 * @since 0.1.0
 * @category models
 */
export interface Workbook {
  /** The format the workbook was read from, e.g. `"csv"`, `"xlsx"`. */
  readonly format: Format
  readonly sheets: ReadonlyArray<SheetInfo>
  /** The limits resolved when the workbook was read; {@link rows} applies its overrides on top. */
  readonly limits: Limits
  /** @internal */
  readonly [RowsSource]: (sheet: SheetInfo, limits: Limits) => Stream.Stream<Row, SpreadsheetError>
}

/**
 * Builds a workbook. For format implementations: pass the limits resolved
 * at read time; `rows` is called with a sheet from `sheets` and the
 * effective limits (those, plus the {@link rows} call's overrides), and
 * must give a stream that can be run again and that uses exactly these
 * limits (not the context's).
 *
 * @since 0.1.0
 * @category constructors
 */
export const make = (options: {
  readonly format: Format
  readonly sheets: ReadonlyArray<SheetInfo>
  readonly limits: Limits
  readonly rows: (sheet: SheetInfo, limits: Limits) => Stream.Stream<Row, SpreadsheetError>
}): Workbook => ({
  format: options.format,
  sheets: options.sheets,
  limits: options.limits,
  [RowsSource]: options.rows
})

/**
 * Finds a sheet by index or by name (exact first, then case-insensitive, as
 * Excel sheet names are).
 *
 * @since 0.1.0
 * @category getters
 */
export const findSheet = (workbook: Workbook, sheet: string | number): SheetInfo | undefined => {
  if (typeof sheet === "number") return workbook.sheets.find((s) => s.index === sheet)
  const exact = workbook.sheets.find((s) => s.name === sheet)
  if (exact !== undefined) return exact
  const lower = sheet.toLowerCase()
  return workbook.sheets.find((s) => s.name.toLowerCase() === lower)
}

/**
 * Gets a sheet by index or name, failing with `SheetNotFound`.
 *
 * @since 0.1.0
 * @category getters
 */
export const sheet = (workbook: Workbook, sheet: string | number): Effect.Effect<SheetInfo, SheetNotFound> => {
  const found = findSheet(workbook, sheet)
  return found === undefined ? Effect.fail(new SheetNotFound({ sheet })) : Effect.succeed(found)
}

/**
 * Streams the rows of one sheet, in order. Rows come in bounded chunks, so
 * the event loop gets control back between them. Hidden rows are included
 * (`hidden: true`). Invalid limit overrides are a defect.
 *
 * @since 0.1.0
 * @category streams
 */
export const rows = (
  workbook: Workbook,
  sheet: string | number,
  options?: RowsOptions
): Stream.Stream<Row, SpreadsheetError> => {
  const found = findSheet(workbook, sheet)
  if (found === undefined) return Stream.fail(new SheetNotFound({ sheet }))
  // suspend: invalid limit overrides (merge throws) die when the stream runs
  return Stream.suspend(() => workbook[RowsSource](found, merge(workbook.limits, options?.limits)))
}

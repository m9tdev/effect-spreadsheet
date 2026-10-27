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
  /** An OLE compound file without a workbook stream: an encrypted OOXML workbook, or another document (a compound file with one is `"xls"`). */
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
 * A defined name (named range, print area, ...), as metadata only: its
 * formula text is not evaluated. `scope` is the index of the sheet it
 * belongs to (in `sheets` order), absent for workbook-wide names.
 *
 * @since 0.1.0
 * @category schemas
 */
export const DefinedName = Schema.Struct({
  name: Schema.String,
  formula: Schema.String,
  scope: Schema.optionalKey(Schema.Int),
  hidden: Schema.Boolean
})

/** @since 0.1.0 @category models */
export type DefinedName = typeof DefinedName.Type

/**
 * A run of columns, 0-based and inclusive.
 *
 * @since 0.1.0
 * @category schemas
 */
export const ColumnSpan = Schema.Struct({ first: Schema.Int, last: Schema.Int })

/** @since 0.1.0 @category models */
export type ColumnSpan = typeof ColumnSpan.Type

/**
 * Sheet layout that isn't part of the rows: hidden columns (sorted,
 * merged spans), merged ranges (as the file lists them) and the declared
 * dimension (informative only, absent if the file has none). From
 * {@link sheetDetails}.
 *
 * @since 0.1.0
 * @category schemas
 */
export const SheetDetails = Schema.Struct({
  hiddenColumns: Schema.Array(ColumnSpan),
  merges: Schema.Array(Range),
  dimension: Schema.optionalKey(Range)
})

/** @since 0.1.0 @category models */
export type SheetDetails = typeof SheetDetails.Type

const NO_DETAILS: SheetDetails = { hiddenColumns: [], merges: [] }

/** @internal */
export const DetailsSource: unique symbol = Symbol.for("effect-spreadsheet/Workbook/details") as never

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
  /** Defined names (metadata only; empty for formats without them). */
  readonly definedNames: ReadonlyArray<DefinedName>
  /** The limits resolved when the workbook was read; {@link rows} applies its overrides on top. */
  readonly limits: Limits
  /** @internal */
  readonly [RowsSource]: (sheet: SheetInfo, limits: Limits) => Stream.Stream<Row, SpreadsheetError>
  /** @internal */
  readonly [DetailsSource]: (sheet: SheetInfo, limits: Limits) => Effect.Effect<SheetDetails, SpreadsheetError>
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
  readonly definedNames?: ReadonlyArray<DefinedName> | undefined
  readonly limits: Limits
  readonly rows: (sheet: SheetInfo, limits: Limits) => Stream.Stream<Row, SpreadsheetError>
  /** Sheet layout; formats without any give no hidden columns, no merges, no dimension. */
  readonly details?: ((sheet: SheetInfo, limits: Limits) => Effect.Effect<SheetDetails, SpreadsheetError>) | undefined
}): Workbook => ({
  format: options.format,
  sheets: options.sheets,
  definedNames: options.definedNames ?? [],
  limits: options.limits,
  [RowsSource]: options.rows,
  [DetailsSource]: options.details ?? (() => Effect.succeed(NO_DETAILS))
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

/**
 * A sheet's hidden columns, merged ranges and declared dimension, read in
 * one bounded pass over the sheet (cell data is skipped, not collected).
 * For formats without them (CSV) all are empty. Uses the workbook's
 * limits, with `options.limits` on top.
 *
 * @since 0.1.0
 * @category getters
 */
export const sheetDetails = (
  workbook: Workbook,
  sheet: string | number,
  options?: RowsOptions
): Effect.Effect<SheetDetails, SpreadsheetError> => {
  const found = findSheet(workbook, sheet)
  if (found === undefined) return Effect.fail(new SheetNotFound({ sheet }))
  return Effect.suspend(() => workbook[DetailsSource](found, merge(workbook.limits, options?.limits)))
}

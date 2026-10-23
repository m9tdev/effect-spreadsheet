/**
 * The format-neutral workbook model: sheets, rows, and a way to stream a
 * sheet's rows.
 *
 * @since 0.1.0
 */
import { Effect, Schema, Stream } from "effect"
import { Cell } from "./Cell.ts"
import { SheetNotFound, type SpreadsheetError } from "./SpreadsheetError.ts"
import type { LimitOverrides } from "./SpreadsheetLimits.ts"

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
 * Options for {@link rows}.
 *
 * @since 0.1.0
 * @category models
 */
export interface RowsOptions {
  /** Overrides on top of the `SpreadsheetLimits` in context. */
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
  readonly format: string
  readonly sheets: ReadonlyArray<SheetInfo>
  /** @internal */
  readonly [RowsSource]: (sheet: SheetInfo, options: RowsOptions) => Stream.Stream<Row, SpreadsheetError>
}

/**
 * Builds a workbook. For format implementations: `rows` is called with a
 * sheet from `sheets` and must give a stream that can be run again.
 *
 * @since 0.1.0
 * @category constructors
 */
export const make = (options: {
  readonly format: string
  readonly sheets: ReadonlyArray<SheetInfo>
  readonly rows: (sheet: SheetInfo, options: RowsOptions) => Stream.Stream<Row, SpreadsheetError>
}): Workbook => ({
  format: options.format,
  sheets: options.sheets,
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
 * (`hidden: true`).
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
  return workbook[RowsSource](found, options ?? {})
}

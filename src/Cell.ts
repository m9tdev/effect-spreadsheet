/**
 * A cell value, independent of the file format it came from.
 *
 * Cells are plain tagged objects (`_tag` is `"Empty"`, `"Text"`, `"Number"`,
 * `"Boolean"`, `"Date"` or `"Error"`) described by Schema, so they can be
 * validated, encoded or sent to a worker. Each cell carries its own 0-based
 * `column`, so rows stay sparse. `formula` (without the leading `=`) and
 * `numFmt` (the number-format code) are present only when the source has
 * them.
 *
 * @since 0.1.0
 */
import { Schema } from "effect"

const meta = {
  /** 0-based column index. */
  column: Schema.Int,
  /** Formula text without the leading `=`, when the file stores one. */
  formula: Schema.optionalKey(Schema.String),
  /** Number-format code (e.g. `"0.00%"`, `"yyyy-mm-dd"`), when the file has one. */
  numFmt: Schema.optionalKey(Schema.String)
}

/**
 * A cell that exists in the file but has no value (e.g. a styled blank).
 *
 * @since 0.1.0
 * @category schemas
 */
export const EmptyCell = Schema.TaggedStruct("Empty", { ...meta })
/** @since 0.1.0 */
export type EmptyCell = typeof EmptyCell.Type

/**
 * @since 0.1.0
 * @category schemas
 */
export const TextCell = Schema.TaggedStruct("Text", { ...meta, value: Schema.String })
/** @since 0.1.0 */
export type TextCell = typeof TextCell.Type

/**
 * @since 0.1.0
 * @category schemas
 */
export const NumberCell = Schema.TaggedStruct("Number", { ...meta, value: Schema.Number })
/** @since 0.1.0 */
export type NumberCell = typeof NumberCell.Type

/**
 * @since 0.1.0
 * @category schemas
 */
export const BooleanCell = Schema.TaggedStruct("Boolean", { ...meta, value: Schema.Boolean })
/** @since 0.1.0 */
export type BooleanCell = typeof BooleanCell.Type

/**
 * A date/time. Spreadsheet dates have no time zone; they are given as the
 * same wall-clock time in UTC.
 *
 * @since 0.1.0
 * @category schemas
 */
export const DateCell = Schema.TaggedStruct("Date", { ...meta, value: Schema.Date })
/** @since 0.1.0 */
export type DateCell = typeof DateCell.Type

/**
 * An error value such as `"#DIV/0!"` or `"#N/A"`.
 *
 * @since 0.1.0
 * @category schemas
 */
export const ErrorCell = Schema.TaggedStruct("Error", { ...meta, value: Schema.String })
/** @since 0.1.0 */
export type ErrorCell = typeof ErrorCell.Type

/**
 * @since 0.1.0
 * @category schemas
 */
export const Cell = Schema.Union([EmptyCell, TextCell, NumberCell, BooleanCell, DateCell, ErrorCell])

/**
 * @since 0.1.0
 * @category models
 */
export type Cell = EmptyCell | TextCell | NumberCell | BooleanCell | DateCell | ErrorCell

/**
 * Optional cell metadata for the constructors.
 *
 * @since 0.1.0
 * @category models
 */
export interface Meta {
  readonly formula?: string | undefined
  readonly numFmt?: string | undefined
}

const withMeta = <C extends Cell>(cell: C, meta: Meta | undefined): C => {
  if (meta !== undefined) {
    const out = cell as { formula?: string; numFmt?: string }
    if (meta.formula !== undefined) out.formula = meta.formula
    if (meta.numFmt !== undefined) out.numFmt = meta.numFmt
  }
  return cell
}

// The constructors build plain objects without validation: readers call them
// once per cell in hot loops.

/** @since 0.1.0 @category constructors */
export const empty = (column: number, meta?: Meta): EmptyCell => withMeta({ _tag: "Empty" as const, column }, meta)

/** @since 0.1.0 @category constructors */
export const text = (column: number, value: string, meta?: Meta): TextCell =>
  withMeta({ _tag: "Text" as const, column, value }, meta)

/** @since 0.1.0 @category constructors */
export const number = (column: number, value: number, meta?: Meta): NumberCell =>
  withMeta({ _tag: "Number" as const, column, value }, meta)

/** @since 0.1.0 @category constructors */
export const boolean = (column: number, value: boolean, meta?: Meta): BooleanCell =>
  withMeta({ _tag: "Boolean" as const, column, value }, meta)

/** @since 0.1.0 @category constructors */
export const date = (column: number, value: Date, meta?: Meta): DateCell =>
  withMeta({ _tag: "Date" as const, column, value }, meta)

/** @since 0.1.0 @category constructors */
export const error = (column: number, value: string, meta?: Meta): ErrorCell =>
  withMeta({ _tag: "Error" as const, column, value }, meta)

/**
 * A cell value: what {@link value} gives for a non-empty cell.
 *
 * @since 0.1.0
 * @category models
 */
export type Value = string | number | boolean | Date

/**
 * The cell's value, `undefined` for `Empty`. Error cells give their code
 * (e.g. `"#N/A"`).
 *
 * @since 0.1.0
 * @category getters
 */
export const value = (cell: Cell): Value | undefined =>
  cell._tag === "Empty" ? undefined : cell.value

/**
 * @since 0.1.0
 * @category guards
 */
export const isEmpty = (cell: Cell): cell is EmptyCell => cell._tag === "Empty"

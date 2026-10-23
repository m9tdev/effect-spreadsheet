/**
 * Typed errors for every reader. Bad input never throws and never becomes a
 * defect: it fails with one of these.
 *
 * @since 0.1.0
 */
import { Schema } from "effect"

/**
 * The bytes are a format this library doesn't read (yet), or no spreadsheet
 * at all. `format` names what was detected, e.g. `"xls"`, `"cfb"`, `"zip"`,
 * `"pdf"`, `"unknown"`.
 *
 * @since 0.1.0
 * @category errors
 */
export class UnsupportedFormat extends Schema.TaggedError<UnsupportedFormat>()("UnsupportedFormat", {
  format: Schema.String,
  reason: Schema.String
}) {
  override get message(): string {
    return `Unsupported format ${this.format}: ${this.reason}`
  }
}

/**
 * The file is recognised but broken: truncated, inconsistent, or violating
 * its format. `part` names the place inside the container (zip entry, CFB
 * stream, ...). `offset` is a byte offset into that part for binary formats;
 * for text formats (CSV) it is a UTF-16 code unit offset into the decoded
 * text.
 *
 * @since 0.1.0
 * @category errors
 */
export class MalformedFile extends Schema.TaggedError<MalformedFile>()("MalformedFile", {
  format: Schema.String,
  part: Schema.optionalKey(Schema.String),
  offset: Schema.optionalKey(Schema.Number),
  reason: Schema.String
}) {
  override get message(): string {
    const where = (this.part !== undefined ? ` in ${this.part}` : "") +
      (this.offset !== undefined ? ` at offset ${this.offset}` : "")
    return `Malformed ${this.format} file${where}: ${this.reason}`
  }
}

/**
 * A {@link SpreadsheetLimits} limit was hit. `limit` is the limit's key
 * (e.g. `"maxColumns"`), `max` its configured value, `actual` the value that
 * exceeded it (it may be a lower bound when counting stopped early).
 *
 * @since 0.1.0
 * @category errors
 */
export class LimitExceeded extends Schema.TaggedError<LimitExceeded>()("LimitExceeded", {
  limit: Schema.String,
  max: Schema.Number,
  actual: Schema.Number
}) {
  override get message(): string {
    return `Limit ${this.limit} exceeded: ${this.actual} > ${this.max}`
  }
}

/**
 * The file is password-protected or encrypted. Decryption is out of scope.
 *
 * @since 0.1.0
 * @category errors
 */
export class Encrypted extends Schema.TaggedError<Encrypted>()("Encrypted", {
  format: Schema.String
}) {
  override get message(): string {
    return `Encrypted ${this.format} file`
  }
}

/**
 * The requested sheet (by name or index) is not in the workbook.
 *
 * @since 0.1.0
 * @category errors
 */
export class SheetNotFound extends Schema.TaggedError<SheetNotFound>()("SheetNotFound", {
  sheet: Schema.Union([Schema.String, Schema.Number])
}) {
  override get message(): string {
    return `Sheet not found: ${typeof this.sheet === "number" ? `#${this.sheet}` : JSON.stringify(this.sheet)}`
  }
}

/**
 * A reader option has a value it can't use, e.g. a delimiter that isn't one
 * character or an encoding label `TextDecoder` doesn't know. `option` names
 * it (e.g. `"delimiter"`, `"encoding"`).
 *
 * @since 0.1.0
 * @category errors
 */
export class InvalidOption extends Schema.TaggedError<InvalidOption>()("InvalidOption", {
  option: Schema.String,
  reason: Schema.String
}) {
  override get message(): string {
    return `Invalid option ${this.option}: ${this.reason}`
  }
}

/**
 * @since 0.1.0
 * @category errors
 */
export type SpreadsheetError =
  | UnsupportedFormat
  | MalformedFile
  | LimitExceeded
  | Encrypted
  | SheetNotFound
  | InvalidOption

/**
 * Schema of the whole union, e.g. to send errors across a worker boundary.
 *
 * @since 0.1.0
 * @category schemas
 */
export const SpreadsheetError = Schema.Union([
  UnsupportedFormat,
  MalformedFile,
  LimitExceeded,
  Encrypted,
  SheetNotFound,
  InvalidOption
])

/**
 * @since 0.1.0
 * @category guards
 */
export const isSpreadsheetError = (u: unknown): u is SpreadsheetError =>
  u instanceof UnsupportedFormat || u instanceof MalformedFile || u instanceof LimitExceeded ||
  u instanceof Encrypted || u instanceof SheetNotFound || u instanceof InvalidOption

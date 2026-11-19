/**
 * Incremental delimited-text parser over UTF-16 text that can be fed in
 * pieces of any size (a piece may end anywhere: inside a quoted field,
 * between `""`, or between CR and LF). Fields come from the lexer the
 * delimiter sniffer shares (`lex.ts`); this builds cells and rows.
 *
 * Behaviour (see README / test/csv for the full list):
 * - Records end at LF, CR or CRLF outside quotes. A trailing line break
 *   does not start a new record; blank lines are empty rows.
 * - A field starting with `"` is quoted: `""` is a literal quote, line
 *   breaks and delimiters are literal. Characters after the closing quote
 *   are appended as-is (`"ab"c` -> `abc`), or rejected in strict mode.
 * - A quote inside an unquoted field is literal.
 * - An unterminated quoted field at end of input keeps what it has, or is
 *   rejected in strict mode.
 * - Empty fields are not emitted as cells (rows are sparse).
 *
 * @internal
 */
import { absurd } from "effect/Function"
import * as Result from "effect/Result"
import * as Cell from "../../Cell.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
import { type FieldEnds, fieldEndsOf, type FieldSink, lexFields, type LexState, Mode } from "./lex.ts"
import { isDelimiter } from "./options.ts"
import { guessDelimiter, SNIFF_CHARS } from "./sniff.ts"

export interface ParserOptions {
  /** The delimiter's character code, or undefined to sniff it. */
  readonly delimiter: number | undefined
  readonly inferTypes: boolean
  readonly strict: boolean
  readonly limits: Limits
}

/**
 * What `CsvParser.feed` does with a piece of text: collect text until the
 * delimiter is known, tokenize it, or nothing more (after the final piece,
 * or after an error, which every later call returns).
 */
type Phase =
  | { readonly _tag: "Sniffing"; readonly pending: string }
  | { readonly _tag: "Parsing"; readonly tokenizer: Tokenizer }
  | { readonly _tag: "Done" }
  | { readonly _tag: "Failed"; readonly error: SpreadsheetError }

export class CsvParser {
  private phase: Phase
  private readonly options: ParserOptions

  constructor(options: ParserOptions) {
    this.options = options
    this.phase = options.delimiter === undefined
      ? { _tag: "Sniffing", pending: "" }
      : { _tag: "Parsing", tokenizer: new Tokenizer(options, options.delimiter, 0) }
  }

  /** The error that stopped parsing, if any. Once set, it stays. */
  get error(): SpreadsheetError | undefined {
    return this.phase._tag === "Failed" ? this.phase.error : undefined
  }

  /**
   * Feeds a piece of text; `final` marks the end of input. Succeeds with the
   * rows completed by this piece, or fails with the error that stopped
   * parsing. If the piece completed rows before the error, those rows come
   * first; the error is then in `error` and is the failure of the next
   * call. After the final piece, later calls give no rows.
   */
  feed(text: string, final: boolean): Result.Result<Array<Row>, SpreadsheetError> {
    const phase = this.phase
    switch (phase._tag) {
      case "Failed":
        return Result.fail(phase.error)
      case "Done":
        return Result.succeed([])
      case "Sniffing": {
        const pending = phase.pending + text
        // collect one character past the sample, so that a sample that is
        // the whole text is known to be whole even before the final piece
        if (pending.length <= SNIFF_CHARS && !final) {
          this.phase = { _tag: "Sniffing", pending }
          return Result.succeed([])
        }
        const tokenizer = this.sniff(pending, final && pending.length <= SNIFF_CHARS)
        this.phase = { _tag: "Parsing", tokenizer }
        return this.step(tokenizer, tokenizer.parse(pending, tokenizer.start), final)
      }
      case "Parsing":
        return this.step(phase.tokenizer, phase.tokenizer.parse(text, 0), final)
      default:
        return absurd(phase)
    }
  }

  /**
   * After a piece was tokenized (`parsed`: the error that stopped it, if
   * any): ends the input on `final`, takes the completed rows and moves to
   * Done or Failed. The tokenizer is not used again after an error.
   */
  private step(
    tokenizer: Tokenizer,
    parsed: SpreadsheetError | undefined,
    final: boolean
  ): Result.Result<Array<Row>, SpreadsheetError> {
    const error = parsed ?? (final ? tokenizer.finish() : undefined)
    const rows = tokenizer.takeRows()
    if (error !== undefined) {
      this.phase = { _tag: "Failed", error }
      return rows.length === 0 ? Result.fail(error) : Result.succeed(rows)
    }
    if (final) this.phase = { _tag: "Done" }
    return Result.succeed(rows)
  }

  /**
   * Decides the delimiter; an Excel `sep=X` line, if present, sets it and is
   * skipped (the tokenizer's `start`). `complete` says that `text` is the
   * whole input. Returns a tokenizer that has parsed nothing yet.
   */
  private sniff(text: string, complete: boolean): Tokenizer {
    const sep = /^sep=(.)(\r\n|\r|\n|$)/.exec(text)
    const char = sep?.[1]
    if (sep !== null && char !== undefined && isDelimiter(char)) {
      return new Tokenizer(this.options, char.charCodeAt(0), sep[0].length)
    }
    return new Tokenizer(this.options, guessDelimiter(text.slice(0, SNIFF_CHARS), complete), 0)
  }
}

/**
 * Splits text into rows once the delimiter is known: a `FieldSink` that
 * builds cells and rows from the shared lexer. Each step returns the error
 * that stops tokenizing, or undefined to go on; after an error the
 * tokenizer is not used again (the parser keeps the error).
 */
class Tokenizer implements FieldSink, LexState {
  private readonly ends: FieldEnds
  /** Characters skipped before the first piece (a `sep=` line). */
  readonly start: number

  // where the lexer stands (LexState): kept on the tokenizer itself, which
  // saves an object per read
  mode: Mode = Mode.FieldStart
  skipLF = false
  recordStarted = false
  private field = ""
  private column = 0
  private cells: Array<Cell.Cell> = []
  private rowIndex = 0
  private cellCount = 0
  /** Characters of the pieces before the current one (for error offsets). */
  private consumed = 0
  private out: Array<Row> = []
  /** The error that stopped the current piece. */
  private error: SpreadsheetError | undefined = undefined

  private readonly options: ParserOptions

  /** `start`: characters of the first piece to skip (a `sep=` line); error offsets count them. */
  constructor(options: ParserOptions, delimiter: number, start: number) {
    this.options = options
    this.ends = fieldEndsOf(delimiter)
    this.start = start
  }

  /** The rows completed since the last call. */
  takeRows(): Array<Row> {
    const rows = this.out
    this.out = []
    return rows
  }

  /** Stops lexing with `error`. */
  private fail(error: SpreadsheetError): false {
    this.error = error
    return false
  }

  text(s: string, from: number, to: number): boolean {
    this.field += s.slice(from, to)
    const max = this.options.limits.maxStringLength
    return this.field.length > max
      ? this.fail(new LimitExceeded({ limit: "maxStringLength", max, actual: this.field.length }))
      : true
  }

  endField(s: string, from: number, to: number): boolean {
    // the field's text so far, then its last part (most fields are only that)
    let value = this.field
    if (to > from) {
      value = value.length === 0 ? s.slice(from, to) : value + s.slice(from, to)
      const max = this.options.limits.maxStringLength
      if (value.length > max) {
        return this.fail(new LimitExceeded({ limit: "maxStringLength", max, actual: value.length }))
      }
    }
    const column = this.column
    const maxColumns = this.options.limits.maxColumns
    if (column >= maxColumns) {
      return this.fail(new LimitExceeded({ limit: "maxColumns", max: maxColumns, actual: column + 1 }))
    }
    this.column = column + 1
    this.field = ""
    // An empty field emits no cell; every emitted cell counts.
    if (value.length === 0) return true
    const maxCells = this.options.limits.maxCellsPerSheet
    if (++this.cellCount > maxCells) {
      return this.fail(new LimitExceeded({ limit: "maxCellsPerSheet", max: maxCells, actual: this.cellCount }))
    }
    this.cells.push(this.options.inferTypes ? infer(column, value) : Cell.text(column, value))
    return true
  }

  endRecord(): boolean {
    const index = this.rowIndex
    const maxRows = this.options.limits.maxRows
    if (index >= maxRows) {
      return this.fail(new LimitExceeded({ limit: "maxRows", max: maxRows, actual: index + 1 }))
    }
    this.out.push({ index, hidden: false, cells: this.cells })
    this.rowIndex = index + 1
    this.cells = []
    this.column = 0
    return true
  }

  afterQuote(at: number): boolean {
    return this.options.strict
      ? this.fail(
        new MalformedFile({
          format: "csv",
          offset: this.consumed + at,
          reason: `unexpected character after closing quote in row ${this.rowIndex + 1}`
        })
      )
      : true
  }

  /**
   * Tokenizes a piece from `from` (the first piece starts after `start`);
   * stops at the first error and returns it.
   */
  parse(s: string, from: number): SpreadsheetError | undefined {
    lexFields(s, from, this, this.ends, this)
    this.consumed += s.length
    return this.error
  }

  /** End of input: ends the last record, if one has started. Returns the error, if any. */
  finish(): SpreadsheetError | undefined {
    if (this.mode === Mode.Quoted && this.options.strict) {
      return new MalformedFile({
        format: "csv",
        offset: this.consumed,
        reason: `unterminated quoted field in row ${this.rowIndex + 1}`
      })
    }
    if (!this.recordStarted) return undefined
    return this.endField("", 0, 0) && this.endRecord() ? undefined : this.error
  }
}

const ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"])

/**
 * Decimal numbers: optional sign, no extra leading zeros in the integer part
 * (so `007` and zip codes stay text), optional fraction and exponent, at
 * most 15 digits before the exponent, not counting zeros before the first
 * non-zero digit (so long IDs keep every digit).
 */
const NUMBER = /^[+-]?(?:(?:0|[1-9]\d*)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/

const significantDigits = (s: string): number => {
  let digits = 0
  let leading = true
  // progress: i increases up to the mantissa end
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    if (c === 0x65 || c === 0x45) break // e / E
    if (c < 0x30 || c > 0x39) continue // sign or dot; the for-loop still advances i
    if (leading && c === 0x30) continue
    leading = false
    digits++
  }
  return digits
}

/** Type inference for `inferTypes: true`. */
export const infer = (column: number, value: string): Cell.Cell => {
  const first = value.charCodeAt(0)
  if ((first >= 0x30 && first <= 0x39) || first === 0x2B || first === 0x2D || first === 0x2E) {
    if (NUMBER.test(value) && significantDigits(value) <= 15) {
      const n = Number(value)
      // `-0`, `-0.0`, `-0e5` are 0 (Excel has no negative zero)
      if (Number.isFinite(n)) return Cell.number(column, n === 0 ? 0 : n)
    }
    return Cell.text(column, value)
  }
  if (value.length === 4 || value.length === 5) {
    const upper = value.toUpperCase()
    if (upper === "TRUE") return Cell.boolean(column, true)
    if (upper === "FALSE") return Cell.boolean(column, false)
  }
  if (first === 0x23 && ERRORS.has(value)) return Cell.error(column, value)
  return Cell.text(column, value)
}

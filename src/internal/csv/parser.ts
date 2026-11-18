/**
 * Incremental delimited-text parser: a state machine over UTF-16 text that
 * can be fed in pieces of any size (a piece may end anywhere: inside a quoted
 * field, between `""`, or between CR and LF).
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
import { isDelimiter } from "./options.ts"
import { guessDelimiter, SNIFF_CHARS } from "./sniff.ts"

const QUOTE = 0x22
const LF = 0x0A
const CR = 0x0D

const State = {
  /** At the start of a field. */
  FieldStart: 0,
  /** Inside an unquoted field. */
  Unquoted: 1,
  /** Inside a quoted field. */
  Quoted: 2,
  /** Just saw a `"` inside a quoted field: `""`, the closing quote, or junk follows. */
  QuoteInQuoted: 3
} as const
type State = typeof State[keyof typeof State]

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
        return this.step(tokenizer, tokenizer.parse(pending.slice(tokenizer.start)), final)
      }
      case "Parsing":
        return this.step(phase.tokenizer, phase.tokenizer.parse(text), final)
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
 * Splits text into rows once the delimiter is known. Each step returns the
 * error that stops tokenizing, or undefined to go on; after an error the
 * tokenizer is not used again (the parser keeps the error).
 */
class Tokenizer {
  private readonly delimiter: number
  /** Characters skipped before the first piece (a `sep=` line). */
  readonly start: number

  private state: State = State.FieldStart
  private field = ""
  /** The current record has consumed at least one character. */
  private recordStarted = false
  /** The previous piece ended with a CR record terminator. */
  private skipLF = false
  private column = 0
  private cells: Array<Cell.Cell> = []
  private rowIndex = 0
  private cellCount = 0
  /** Characters consumed before the current piece (for error offsets). */
  private consumed = 0
  private out: Array<Row> = []

  private readonly options: ParserOptions

  /** `start`: characters skipped before the first piece (a `sep=` line); error offsets count them. */
  constructor(options: ParserOptions, delimiter: number, start: number) {
    this.options = options
    this.delimiter = delimiter
    this.start = start
    this.consumed = start
  }

  /** The rows completed since the last call. */
  takeRows(): Array<Row> {
    const rows = this.out
    this.out = []
    return rows
  }

  private appendField(s: string, from: number, to: number): LimitExceeded | undefined {
    if (to > from) {
      this.field += s.slice(from, to)
      const max = this.options.limits.maxStringLength
      if (this.field.length > max) {
        return new LimitExceeded({ limit: "maxStringLength", max, actual: this.field.length })
      }
    }
    return undefined
  }

  private endField(): LimitExceeded | undefined {
    const column = this.column
    const maxColumns = this.options.limits.maxColumns
    if (column >= maxColumns) {
      return new LimitExceeded({ limit: "maxColumns", max: maxColumns, actual: column + 1 })
    }
    this.column = column + 1
    const value = this.field
    this.field = ""
    this.state = State.FieldStart
    // An empty field emits no cell; every emitted cell counts.
    if (value.length === 0) return undefined
    const maxCells = this.options.limits.maxCellsPerSheet
    if (++this.cellCount > maxCells) {
      return new LimitExceeded({ limit: "maxCellsPerSheet", max: maxCells, actual: this.cellCount })
    }
    this.cells.push(this.options.inferTypes ? infer(column, value) : Cell.text(column, value))
    return undefined
  }

  private endRecord(): LimitExceeded | undefined {
    const index = this.rowIndex
    const maxRows = this.options.limits.maxRows
    if (index >= maxRows) {
      return new LimitExceeded({ limit: "maxRows", max: maxRows, actual: index + 1 })
    }
    this.out.push({ index, hidden: false, cells: this.cells })
    this.rowIndex = index + 1
    this.cells = []
    this.column = 0
    this.recordStarted = false
    return undefined
  }

  /** Ends the field at a delimiter or line break `c`. */
  private terminate(c: number): LimitExceeded | undefined {
    const error = this.endField()
    if (error !== undefined || c === this.delimiter) return error
    if (c === CR) this.skipLF = true
    return this.endRecord()
  }

  /** Tokenizes a piece; stops at the first error and returns it. */
  parse(s: string): SpreadsheetError | undefined {
    const n = s.length
    const delim = this.delimiter
    let error: SpreadsheetError | undefined
    let i = 0
    // progress: every branch below moves `i` forward by at least 1, or ends
    // the loop (error). `i` never decreases.
    while (i < n && error === undefined) {
      if (this.skipLF) {
        this.skipLF = false
        if (s.charCodeAt(i) === LF) {
          i++
          continue // `i` advanced
        }
      }
      switch (this.state) {
        case State.FieldStart:
        case State.Unquoted: {
          if (this.state === State.FieldStart && s.charCodeAt(i) === QUOTE) {
            this.state = State.Quoted
            this.recordStarted = true
            i++
            break
          }
          this.state = State.Unquoted
          this.recordStarted = true
          let j = i
          // progress: j increases up to n
          for (; j < n; j++) {
            const c = s.charCodeAt(j)
            if (c === delim || c === LF || c === CR) break
          }
          error = this.appendField(s, i, j)
          if (error === undefined && j < n) error = this.terminate(s.charCodeAt(j))
          i = j + 1 // > i; when j === n this ends the loop
          break
        }
        case State.Quoted: {
          const j = s.indexOf('"', i)
          if (j < 0) {
            error = this.appendField(s, i, n)
            i = n
          } else {
            error = this.appendField(s, i, j)
            this.state = State.QuoteInQuoted
            i = j + 1
          }
          break
        }
        case State.QuoteInQuoted: {
          const c = s.charCodeAt(i)
          if (c === QUOTE) {
            error = this.appendField('"', 0, 1)
            this.state = State.Quoted
            i++
          } else if (c === delim || c === LF || c === CR) {
            error = this.terminate(c)
            i++
          } else if (this.options.strict) {
            error = new MalformedFile({
              format: "csv",
              offset: this.consumed + i,
              reason: `unexpected character after closing quote in row ${this.rowIndex + 1}`
            })
          } else {
            // Lenient: the closing quote ended the quoted part; the rest of
            // the field is literal. `c` is not a terminator, so the next
            // iteration (Unquoted) consumes at least `c`.
            this.state = State.Unquoted
            let j = i + 1
            // progress: j increases up to n
            for (; j < n; j++) {
              const d = s.charCodeAt(j)
              if (d === delim || d === LF || d === CR) break
            }
            error = this.appendField(s, i, j)
            if (error === undefined && j < n) error = this.terminate(s.charCodeAt(j))
            i = j + 1
          }
          break
        }
      }
    }
    this.consumed += n
    return error
  }

  /** End of input: ends the last record, if one has started. Returns the error, if any. */
  finish(): SpreadsheetError | undefined {
    if (this.state === State.Quoted && this.options.strict) {
      return new MalformedFile({
        format: "csv",
        offset: this.consumed,
        reason: `unterminated quoted field in row ${this.rowIndex + 1}`
      })
    }
    if (!this.recordStarted) return undefined
    return this.endField() ?? this.endRecord()
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

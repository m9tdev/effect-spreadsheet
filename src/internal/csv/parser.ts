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
import * as Cell from "../../Cell.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
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
  /** Delimiter character, or undefined to sniff it. */
  readonly delimiter: string | undefined
  readonly inferTypes: boolean
  readonly strict: boolean
  readonly limits: Limits
}

export class CsvParser {
  private delimiter: number
  private sniffing: boolean
  private pending = ""

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
  private failure: SpreadsheetError | undefined

  private readonly options: ParserOptions

  constructor(options: ParserOptions) {
    this.options = options
    this.sniffing = options.delimiter === undefined
    this.delimiter = options.delimiter === undefined ? 0x2C : options.delimiter.charCodeAt(0)
  }

  /** The delimiter in use (after sniffing). */
  get delimiterChar(): string {
    return String.fromCharCode(this.delimiter)
  }

  /**
   * Feeds a piece of text; `final` marks the end of input. Returns the rows
   * completed by this piece, or the error that stopped parsing.
   */
  feed(text: string, final: boolean): Array<Row> | SpreadsheetError {
    if (this.failure !== undefined) return this.failure
    if (this.sniffing) {
      this.pending += text
      if (this.pending.length < SNIFF_CHARS && !final) return []
      this.sniffing = false
      text = this.pending
      this.pending = ""
      text = this.sniff(text)
    }
    this.parse(text)
    if (final && this.failure === undefined) this.finish()
    if (this.failure !== undefined) return this.failure
    const rows = this.out
    this.out = []
    return rows
  }

  /** Decides the delimiter; consumes an Excel `sep=X` line if present. */
  private sniff(text: string): string {
    const sep = /^sep=(.)(\r\n|\r|\n|$)/.exec(text)
    if (sep !== null && sep[1] !== '"' && sep[1] !== "\r" && sep[1] !== "\n") {
      this.delimiter = sep[1]!.charCodeAt(0)
      this.consumed += sep[0].length
      return text.slice(sep[0].length)
    }
    this.delimiter = guessDelimiter(text.slice(0, SNIFF_CHARS))
    return text
  }

  private fail(error: SpreadsheetError): void {
    if (this.failure === undefined) this.failure = error
  }

  private appendField(s: string, from: number, to: number): void {
    if (to > from) {
      this.field += s.slice(from, to)
      const max = this.options.limits.maxStringLength
      if (this.field.length > max) {
        this.fail(new LimitExceeded({ limit: "maxStringLength", max, actual: this.field.length }))
      }
    }
  }

  private endField(): void {
    const column = this.column
    const maxColumns = this.options.limits.maxColumns
    if (column >= maxColumns) {
      this.fail(new LimitExceeded({ limit: "maxColumns", max: maxColumns, actual: column + 1 }))
      return
    }
    this.column = column + 1
    const value = this.field
    this.field = ""
    this.state = State.FieldStart
    // An empty field emits no cell; every emitted cell counts.
    if (value.length === 0) return
    const maxCells = this.options.limits.maxCellsPerSheet
    if (++this.cellCount > maxCells) {
      this.fail(new LimitExceeded({ limit: "maxCellsPerSheet", max: maxCells, actual: this.cellCount }))
      return
    }
    this.cells.push(this.options.inferTypes ? infer(column, value) : Cell.text(column, value))
  }

  private endRecord(): void {
    const index = this.rowIndex
    const maxRows = this.options.limits.maxRows
    if (index >= maxRows) {
      this.fail(new LimitExceeded({ limit: "maxRows", max: maxRows, actual: index + 1 }))
      return
    }
    this.out.push({ index, hidden: false, cells: this.cells })
    this.rowIndex = index + 1
    this.cells = []
    this.column = 0
    this.recordStarted = false
  }

  /** Ends the field at a delimiter or line break `c` found at `s[j]`. */
  private terminate(c: number): void {
    this.endField()
    if (c !== this.delimiter) {
      if (this.failure === undefined) this.endRecord()
      if (c === CR) this.skipLF = true
    }
  }

  private parse(s: string): void {
    const n = s.length
    const delim = this.delimiter
    let i = 0
    // progress: every branch below moves `i` forward by at least 1, or ends
    // the loop (failure). `i` never decreases.
    while (i < n && this.failure === undefined) {
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
          this.appendField(s, i, j)
          if (j < n) this.terminate(s.charCodeAt(j))
          i = j + 1 // > i; when j === n this ends the loop
          break
        }
        case State.Quoted: {
          const j = s.indexOf('"', i)
          if (j < 0) {
            this.appendField(s, i, n)
            i = n
          } else {
            this.appendField(s, i, j)
            this.state = State.QuoteInQuoted
            i = j + 1
          }
          break
        }
        case State.QuoteInQuoted: {
          const c = s.charCodeAt(i)
          if (c === QUOTE) {
            this.appendField('"', 0, 1)
            this.state = State.Quoted
            i++
          } else if (c === delim || c === LF || c === CR) {
            this.terminate(c)
            i++
          } else if (this.options.strict) {
            this.fail(
              new MalformedFile({
                format: "csv",
                offset: this.consumed + i,
                reason: `unexpected character after closing quote in row ${this.rowIndex + 1}`
              })
            )
            i++
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
            this.appendField(s, i, j)
            if (j < n) this.terminate(s.charCodeAt(j))
            i = j + 1
          }
          break
        }
      }
    }
    this.consumed += n
  }

  private finish(): void {
    if (this.state === State.Quoted) {
      if (this.options.strict) {
        this.fail(
          new MalformedFile({
            format: "csv",
            offset: this.consumed,
            reason: `unterminated quoted field in row ${this.rowIndex + 1}`
          })
        )
        return
      }
    }
    if (this.recordStarted) {
      this.endField()
      if (this.failure === undefined) this.endRecord()
    }
  }
}

const ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"])

/**
 * Decimal numbers: optional sign, no leading zeros (so `007` and zip codes
 * stay text), optional fraction and exponent, at most 15 digits not
 * counting leading zeros (so long IDs keep every digit).
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

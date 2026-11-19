/**
 * Field lexing for delimited text, shared by the parser and the delimiter
 * sniffer: one state machine for quotes and line breaks, with the delimiters
 * given as a set (the one delimiter when parsing, every candidate while
 * sniffing). It can be fed in pieces of any size: a piece may end anywhere,
 * inside a quoted field, between `""`, or between CR and LF.
 *
 * - Records end at LF, CR or CRLF outside quotes.
 * - A field starting with `"` is quoted: `""` is a literal quote, line
 *   breaks and delimiters are literal. After the closing quote comes a
 *   delimiter, a line break, or other text (see `FieldSink.afterQuote`).
 * - A quote inside an unquoted field is literal.
 *
 * @internal
 */

const QUOTE = 0x22
const LF = 0x0A
const CR = 0x0D

/**
 * Which characters end a field, as a table indexed by UTF-16 code unit: 1
 * at each delimiter and at CR and LF, 0 elsewhere. Built by `fieldEnds`.
 * A table rather than a predicate function: the lexer's inner loop looks
 * up every character, and a lookup keeps it as fast as comparing with one
 * known delimiter, whether the lexer serves the parser or the sniffer.
 */
export type FieldEnds = Uint8Array

/**
 * The field ends for a set of delimiters: character codes other than a
 * quote, CR or LF (the lexer gives those their own meaning).
 */
export const fieldEnds = (delimiters: ReadonlyArray<number>): FieldEnds => {
  const ends = new Uint8Array(0x10000)
  ends[LF] = 1
  ends[CR] = 1
  // progress: one delimiter per iteration
  for (const d of delimiters) ends[d] = 1
  return ends
}

export const Mode = {
  /** At the start of a field. */
  FieldStart: 0,
  /** Inside an unquoted field. */
  Unquoted: 1,
  /** Inside a quoted field. */
  Quoted: 2,
  /** Just saw a `"` inside a quoted field: `""`, the closing quote, or other text follows. */
  QuoteInQuoted: 3
} as const
export type Mode = typeof Mode[keyof typeof Mode]

/** Where lexing stands between pieces of text. */
export interface LexState {
  mode: Mode
  /** The previous piece ended with a CR record terminator: a LF next belongs to it. */
  skipLF: boolean
  /** The current record has consumed a character other than its line break. */
  recordStarted: boolean
}

export const lexState = (): LexState => ({ mode: Mode.FieldStart, skipLF: false, recordStarted: false })

/**
 * What the lexer reports, in order. A field's characters come in parts:
 * `text` for each part but the last, and the last part (maybe empty) with
 * the field's end, so an unquoted field within one piece is one call.
 * Each method returns true to go on, or false to stop lexing (the sink
 * keeps why); after false the lexer returns at once and the state is not
 * meant to be lexed further.
 */
export interface FieldSink {
  /** Characters `from` to `to` (exclusive, `to > from`) of `s` belong to the current field. */
  text(s: string, from: number, to: number): boolean
  /**
   * Characters `from` to `to` (exclusive, maybe none) of `s` end the
   * current field, and the character at `to` ends it: a delimiter, or a
   * line break (then `endRecord` follows).
   */
  endField(s: string, from: number, to: number): boolean
  /**
   * The record ends at the line break that ended its last field.
   * `started`: the record had a character before its line break.
   */
  endRecord(started: boolean): boolean
  /**
   * The character at `at` of the piece follows a closing quote and is not a
   * quote, delimiter or line break. True reads it leniently: the quoted
   * part has ended and the rest of the field, from `at`, is literal.
   */
  afterQuote(at: number): boolean
}

/**
 * Lexes `s` from `from`, reporting to `sink` and keeping its place in
 * `state` for the next piece. Stops at the end of `s`, or as soon as the
 * sink returns false.
 */
export const lexFields = (s: string, from: number, state: LexState, ends: FieldEnds, sink: FieldSink): void => {
  const n = s.length
  let i = from
  // the state in locals while lexing (the hot loop), written back at the end
  let mode: Mode = state.mode
  let started = state.recordStarted
  let skipLF = state.skipLF
  if (skipLF && i < n) {
    skipLF = false
    if (s.charCodeAt(i) === LF) i++
  }
  // progress: every iteration moves `i` forward by at least 1, or leaves
  // the loop. `i` never decreases.
  while (i < n) {
    if (mode === Mode.Quoted) {
      const j = s.indexOf('"', i)
      if (j < 0) {
        // i < n: the rest of the piece is in the field
        sink.text(s, i, n)
        i = n
        break
      }
      mode = Mode.QuoteInQuoted
      if (j > i && !sink.text(s, i, j)) break
      i = j + 1
      continue // `i` advanced
    }
    // the field's last part starts here
    const last = i
    if (mode === Mode.QuoteInQuoted) {
      const c = s.charCodeAt(i)
      if (c === QUOTE) {
        mode = Mode.Quoted
        if (!sink.text('"', 0, 1)) break
        i++
        continue // `i` advanced
      }
      if (ends[c] === 0) {
        // text after the closing quote: the scan below consumes at least `c`
        if (!sink.afterQuote(i)) break
        mode = Mode.Unquoted
      }
    } else if (mode === Mode.FieldStart && s.charCodeAt(i) === QUOTE) {
      mode = Mode.Quoted
      started = true
      i++
      continue // `i` advanced
    }
    if (mode !== Mode.QuoteInQuoted) {
      mode = Mode.Unquoted
      let j = i
      // progress: j increases up to n
      for (; j < n; j++) {
        if (ends[s.charCodeAt(j)] !== 0) break
      }
      if (j === n) {
        // the field goes on in the next piece
        if (j > i) {
          started = true
          sink.text(s, i, j)
        }
        i = n
        break
      }
      i = j
    }
    // the character at `i` ends the field: a delimiter or a line break
    const c = s.charCodeAt(i)
    mode = Mode.FieldStart
    if (!sink.endField(s, last, i)) break
    if (c !== LF && c !== CR) {
      started = true
      i++
      continue // `i` advanced
    }
    const had = started || i > last
    started = false
    if (!sink.endRecord(had)) break
    i++
    if (c === CR) {
      // a LF right after belongs to this line break, also in the next piece
      if (i === n) skipLF = true
      else if (s.charCodeAt(i) === LF) i++
    }
  }
  state.mode = mode
  state.recordStarted = started
  state.skipLF = skipLF
}

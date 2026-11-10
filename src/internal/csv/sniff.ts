/**
 * Delimiter sniffing.
 *
 * @internal
 */
import { absurd } from "effect/Function"

/** Characters of text looked at to guess the delimiter. */
export const SNIFF_CHARS = 16 * 1024

/** Candidates, in tie-break order: comma, tab, semicolon, pipe. */
const CANDIDATES: ReadonlyArray<number> = [0x2C, 0x09, 0x3B, 0x7C]

const QUOTE = 0x22
const LF = 0x0A
const CR = 0x0D

const MAX_LINES = 64

/**
 * Where the scan is within a field, with the parser's quote rules: a `"`
 * opens a quoted field only at the start of a field, and is literal anywhere
 * else.
 */
type Mode = "FieldStart" | "Unquoted" | "Quoted" | "QuoteInQuoted"

/** How `sample` splits into lines and fields when `delim` is the delimiter. */
interface Reading {
  /** Delimiters on each non-blank line (at most `MAX_LINES` lines). */
  readonly lines: Array<number>
  /** Delimiters seen, including on a last line cut off by the sample's end. */
  readonly total: number
}

/**
 * Reads `sample` as the parser would with delimiter `delim`, counting the
 * delimiters outside quoted fields on each line.
 */
const read = (sample: string, delim: number): Reading => {
  const lines: Array<number> = []
  let total = 0
  let current = 0
  let lineHasText = false
  let mode: Mode = "FieldStart"
  const n = sample.length
  // progress: i increases up to n (the skips below only move it forward);
  // lines are capped at MAX_LINES
  for (let i = 0; i < n && lines.length < MAX_LINES; i++) {
    if (mode === "Quoted") {
      // line breaks and delimiters are part of the field: skip to its next
      // quote (the opening quote already marked the line as non-blank)
      const j = sample.indexOf('"', i)
      if (j < 0) break
      i = j
      mode = "QuoteInQuoted"
      continue
    }
    const c = sample.charCodeAt(i)
    if (c === LF || c === CR) {
      if (lineHasText) lines.push(current)
      current = 0
      lineHasText = false
      mode = "FieldStart"
      continue
    }
    lineHasText = true
    if (c === delim) {
      current++
      total++
      mode = "FieldStart"
      continue
    }
    switch (mode) {
      case "FieldStart":
        mode = c === QUOTE ? "Quoted" : "Unquoted"
        break
      case "Unquoted":
        break
      case "QuoteInQuoted":
        // `""` is a literal quote; anything else follows the closing quote
        // and is literal (the parser's lenient reading)
        mode = c === QUOTE ? "Quoted" : "Unquoted"
        break
      default:
        absurd(mode)
    }
    if (mode === "Unquoted") {
      // the rest of the field is literal: skip to its end
      while (i + 1 < n) {
        const d = sample.charCodeAt(i + 1)
        if (d === delim || d === LF || d === CR) break
        i++
      }
    }
  }
  // A last line without a line break counts only if the sample is the whole
  // text (it could be cut otherwise).
  if (lineHasText && n < SNIFF_CHARS && lines.length < MAX_LINES) lines.push(current)
  return { lines, total }
}

/**
 * Guesses the delimiter of `sample`. Each candidate is tried as the
 * delimiter, splitting lines and fields as the parser would with it, so
 * quoted fields hide that candidate and line breaks inside them.
 *
 * 1. A candidate that occurs the same non-zero number of times on every
 *    line (at least 2 non-blank lines) wins; among several, the one with
 *    more occurrences per line, then the tie-break order.
 * 2. Otherwise the most frequent candidate wins, then the
 *    tie-break order; with none at all, `,`.
 */
export const guessDelimiter = (sample: string): number => {
  const readings = CANDIDATES.map((delim) => read(sample, delim))
  let best = -1
  let bestPerLine = 0
  readings.forEach(({ lines }, t) => {
    if (lines.length < 2) return
    const first = lines[0]!
    if (first > bestPerLine && lines.every((count) => count === first)) {
      best = t
      bestPerLine = first
    }
  })
  if (best >= 0) return CANDIDATES[best]!
  let bestTotal = 0
  readings.forEach(({ total }, t) => {
    if (total > bestTotal) {
      best = t
      bestTotal = total
    }
  })
  return best >= 0 ? CANDIDATES[best]! : 0x2C
}

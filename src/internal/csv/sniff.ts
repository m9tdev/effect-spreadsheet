/**
 * Delimiter sniffing.
 *
 * @internal
 */
import { fieldEnds, type FieldSink, lexFields, lexState } from "./lex.ts"

/** Characters of text looked at to guess the delimiter. */
export const SNIFF_CHARS = 16 * 1024

/** Candidates, in tie-break order: comma, tab, semicolon, pipe. */
const CANDIDATES = [0x2C, 0x09, 0x3B, 0x7C] as const
const K = CANDIDATES.length

const CANDIDATE_ENDS = fieldEnds(CANDIDATES)

/** A field end's index in CANDIDATES, by character code (all below 128), or -1. */
const CANDIDATE_INDEX = new Int8Array(128).fill(-1)
// progress: t increases up to K
for (let t = 0; t < K; t++) CANDIDATE_INDEX[CANDIDATES[t]!] = t

const MAX_LINES = 64

/**
 * Counts the candidates that end fields, per line, with every candidate
 * taken as a delimiter (a field starts at the start of a line and after any
 * candidate, so quotes follow the parser's rules). Lines without a
 * character are skipped; it stops after MAX_LINES lines.
 */
class CandidateCounts implements FieldSink {
  /**
   * The counts of line `l` at `l * K` (by candidate), for the `lines` lines
   * read and then the line being read. A line without a character has no
   * candidate, so its slot is still zero for the next line.
   */
  readonly counts: Array<number> = new Array<number>((MAX_LINES + 1) * K).fill(0)
  lines = 0

  text(): boolean {
    return true
  }

  endField(s: string, _from: number, to: number): boolean {
    // a candidate, or a line break
    const t = CANDIDATE_INDEX[s.charCodeAt(to)]!
    if (t >= 0) this.counts[this.lines * K + t]!++
    return true
  }

  endRecord(started: boolean): boolean {
    if (started) this.lines++
    return this.lines < MAX_LINES
  }

  afterQuote(): boolean {
    // the parser's lenient reading: the rest of the field is literal
    return true
  }
}

/**
 * Guesses the delimiter of `sample`. Candidates inside quoted fields are
 * skipped. `complete` says that the sample is the whole text, so a last line
 * without a line break is a whole line; otherwise the sample may end inside
 * that line, which is then left out.
 *
 * 1. A candidate that occurs the same non-zero number of times on every
 *    line (at least 2 non-blank lines) wins; among several, the one with
 *    more occurrences per line, then the tie-break order.
 * 2. Otherwise the most frequent candidate wins, then the
 *    tie-break order; with none at all, `,`.
 */
export const guessDelimiter = (sample: string, complete: boolean): number => {
  const state = lexState()
  const sink = new CandidateCounts()
  lexFields(sample, 0, state, CANDIDATE_ENDS, sink)
  const counts = sink.counts
  // the line being read counts as a line only if the sample is whole
  const lines = state.recordStarted && complete && sink.lines < MAX_LINES ? sink.lines + 1 : sink.lines

  let best = -1
  let bestPerLine = 0
  if (lines >= 2) {
    for (let t = 0; t < K; t++) {
      const first = counts[t]!
      if (first === 0) continue
      let l = 1
      // progress: l increases up to lines
      while (l < lines && counts[l * K + t] === first) l++
      if (l < lines) continue
      if (first > bestPerLine) {
        best = t
        bestPerLine = first
      }
    }
  }
  if (best >= 0) return CANDIDATES[best]!
  let bestTotal = 0
  for (let t = 0; t < K; t++) {
    // every line read, and the line being read (zero if lexing stopped after MAX_LINES)
    let total = 0
    // progress: l increases up to sink.lines
    for (let l = 0; l <= sink.lines; l++) total += counts[l * K + t]!
    if (total > bestTotal) {
      best = t
      bestTotal = total
    }
  }
  return best >= 0 ? CANDIDATES[best]! : 0x2C
}

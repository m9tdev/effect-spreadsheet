/**
 * Delimiter sniffing.
 *
 * @internal
 */

/** Characters of text looked at to guess the delimiter. */
export const SNIFF_CHARS = 16 * 1024

/** Candidates, in tie-break order: comma, tab, semicolon, pipe. */
const CANDIDATES = [0x2C, 0x09, 0x3B, 0x7C] as const

const MAX_LINES = 64

/**
 * Guesses the delimiter of `sample`. Quotes are skipped.
 *
 * 1. A candidate that occurs the same non-zero number of times on every
 *    line (at least 2 non-blank lines) wins; among several, the one with
 *    more occurrences per line, then the tie-break order.
 * 2. Otherwise the most frequent candidate wins, then the
 *    tie-break order; with none at all, `,`.
 */
export const guessDelimiter = (sample: string): number => {
  const k = CANDIDATES.length
  const totals = [0, 0, 0, 0]
  const lineCounts: Array<Array<number>> = []
  let current = [0, 0, 0, 0]
  let lineHasText = false
  let inQuotes = false
  const n = sample.length
  // progress: i increases up to n; lines are capped at MAX_LINES
  for (let i = 0; i < n && lineCounts.length < MAX_LINES; i++) {
    const c = sample.charCodeAt(i)
    if (c === 0x22) {
      inQuotes = !inQuotes
      lineHasText = true
    } else if (!inQuotes && (c === 0x0A || c === 0x0D)) {
      if (lineHasText) lineCounts.push(current)
      current = [0, 0, 0, 0]
      lineHasText = false
    } else {
      lineHasText = true
      if (!inQuotes) {
        for (let t = 0; t < k; t++) {
          if (c === CANDIDATES[t]) {
            current[t]!++
            totals[t]!++
          }
        }
      }
    }
  }
  // A last line without a line break counts only if the sample is the whole
  // text (it could be cut otherwise).
  if (lineHasText && n < SNIFF_CHARS && lineCounts.length < MAX_LINES) lineCounts.push(current)

  let best = -1
  let bestPerLine = 0
  if (lineCounts.length >= 2) {
    for (let t = 0; t < k; t++) {
      const first = lineCounts[0]![t]!
      if (first === 0) continue
      if (!lineCounts.every((line) => line[t] === first)) continue
      if (first > bestPerLine) {
        best = t
        bestPerLine = first
      }
    }
  }
  if (best >= 0) return CANDIDATES[best]!
  let bestTotal = 0
  for (let t = 0; t < k; t++) {
    if (totals[t]! > bestTotal) {
      best = t
      bestTotal = totals[t]!
    }
  }
  return best >= 0 ? CANDIDATES[best]! : 0x2C
}

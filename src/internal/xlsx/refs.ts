/**
 * A1-style cell references and ranges.
 *
 * @internal
 */

/** A parsed reference: 0-based row and column, either of which may be absent (`A`, `1`). */
export interface CellRef {
  readonly row: number | undefined
  readonly column: number | undefined
}

/** Column letters past this length are certainly beyond any column limit (26^7 > 2^32). */
const MAX_LETTERS = 7
/** Row digits past this length are certainly beyond any row limit. */
const MAX_DIGITS = 10

/**
 * Parses `A1`, `$A$1`, `XFD1048576`, also just letters or just digits.
 * Values too large to be real (more than 7 letters or 10 digits) come back
 * as `Infinity`, for the caller's limit check. `undefined` if it isn't a
 * reference.
 */
export const parseCellRef = (ref: string): CellRef | undefined => {
  let i = 0
  const n = ref.length
  if (ref.charCodeAt(i) === 0x24) i++ // $
  let column = 0
  let letters = 0
  // progress: i increases up to n
  while (i < n) {
    const c = ref.charCodeAt(i) | 0x20 // lower case
    if (c < 0x61 || c > 0x7A) break
    if (letters < MAX_LETTERS) column = column * 26 + (c - 0x60)
    letters++
    i++
  }
  if (i < n && ref.charCodeAt(i) === 0x24) i++ // $
  let row = 0
  let digits = 0
  // progress: i increases up to n
  while (i < n) {
    const c = ref.charCodeAt(i)
    if (c < 0x30 || c > 0x39) break
    if (digits < MAX_DIGITS) row = row * 10 + (c - 0x30)
    digits++
    i++
  }
  if (i !== n || (letters === 0 && digits === 0)) return undefined
  if (digits > 0 && row === 0 && digits <= MAX_DIGITS) return undefined // rows start at 1
  return {
    column: letters === 0 ? undefined : letters > MAX_LETTERS ? Infinity : column - 1,
    row: digits === 0 ? undefined : digits > MAX_DIGITS ? Infinity : row - 1
  }
}

/** A 0-based inclusive range. */
export interface ParsedRange {
  readonly firstRow: number
  readonly firstColumn: number
  readonly lastRow: number
  readonly lastColumn: number
}

/** Parses `A1:B2` or a single cell `A1` (as a one-cell range); `undefined` if invalid. */
export const parseRange = (ref: string): ParsedRange | undefined => {
  const colon = ref.indexOf(":")
  const a = parseCellRef(colon < 0 ? ref : ref.slice(0, colon))
  const b = colon < 0 ? a : parseCellRef(ref.slice(colon + 1))
  if (a === undefined || b === undefined) return undefined
  if (a.row === undefined || a.column === undefined || b.row === undefined || b.column === undefined) return undefined
  if (![a.row, a.column, b.row, b.column].every(Number.isFinite)) return undefined
  return {
    firstRow: Math.min(a.row, b.row),
    firstColumn: Math.min(a.column, b.column),
    lastRow: Math.max(a.row, b.row),
    lastColumn: Math.max(a.column, b.column)
  }
}

/** Column index (0-based) to letters. */
export const columnName = (column: number): string => {
  let out = ""
  let c = column + 1
  // progress: c shrinks by a factor of 26
  while (c > 0) {
    const r = (c - 1) % 26
    out = String.fromCharCode(65 + r) + out
    c = Math.floor((c - 1) / 26)
  }
  return out
}

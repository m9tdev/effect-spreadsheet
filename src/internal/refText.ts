/**
 * Reference text as Excel writes it in XLSX formulas (`Sheet!$A$1`,
 * `'My Sheet'!$1:$2`, `[1]Data!$A:$A`), for the defined names of the
 * binary formats (XLS, XLSB), whose formulas are parsed tokens.
 *
 * Both formats give a reference's column as a 16-bit word with flags: bit
 * 14 for a relative column, bit 15 for a relative row, the column below.
 *
 * @internal
 */
import { columnName } from "./xlsx/refs.ts"

/** The column number of letters (`A` 1, `XFD` 16384). */
const columnNumber = (letters: string): number => {
  let n = 0
  // progress: k up to letters.length
  for (let k = 0; k < letters.length; k++) n = n * 26 + (letters.toUpperCase().charCodeAt(k) - 64)
  return n
}

/**
 * Whether a sheet name needs quotes in a reference, as Excel writes them:
 * when it isn't name-like (a letter or `_` first, then letters, digits, `_`
 * and `.`; characters beyond ASCII count as letters), or when it reads as
 * a cell reference (`A1`, up to column XFD) or an R1C1 one (`R`, `C2`,
 * `R1C1`).
 */
export const needsQuotes = (name: string): boolean => {
  if (!/^[A-Za-z_\u0080-\uFFFF][A-Za-z0-9_.\u0080-\uFFFF]*$/.test(name)) return true
  if (/^(?:R\d*|C\d*|R\d*C\d*)$/i.test(name)) return true
  const a1 = /^([A-Za-z]{1,3})\d+$/.exec(name)
  return a1 !== null && columnNumber(a1[1]!) <= 16384
}

/** A sheet prefix (`Sheet!`, `'My Sheet'!`, `Jan:Mar!`, `[1]Data!`): quoted as a whole when any sheet name needs it. */
export const sheetPrefix = (book: string, names: ReadonlyArray<string>): string => {
  const text = book + names.join(":")
  return names.some(needsQuotes) ? `'${text.replace(/'/g, "''")}'!` : `${text}!`
}

/** The last row and column of a format's sheets (0-based): whole rows and columns are written as `$1:$2`, `$A:$B`. */
export interface Grid {
  readonly lastRow: number
  readonly lastColumn: number
}

export const BIFF8_GRID: Grid = { lastRow: 0xFFFF, lastColumn: 0xFF }
export const BIFF12_GRID: Grid = { lastRow: 0xFFFFF, lastColumn: 0x3FFF }

const colText = (col: number): string => `${(col & 0x4000) !== 0 ? "" : "$"}${columnName(col & 0x3FFF)}`
const rowText = (row: number, col: number): string => `${(col & 0x8000) !== 0 ? "" : "$"}${row + 1}`

/** A cell: `$A$1` (`A1` for relative parts). */
export const cellText = (row: number, col: number): string => colText(col) + rowText(row, col)

/** An area: whole rows as `$1:$2`, whole columns as `$A:$B`, else `$A$1:$B$2`. */
export const areaText = (r1: number, r2: number, c1: number, c2: number, grid: Grid): string => {
  if ((c1 & 0x3FFF) === 0 && (c2 & 0x3FFF) === grid.lastColumn) return `${rowText(r1, c1)}:${rowText(r2, c2)}`
  if (r1 === 0 && r2 === grid.lastRow) return `${colText(c1)}:${colText(c2)}`
  return `${cellText(r1, c1)}:${cellText(r2, c2)}`
}

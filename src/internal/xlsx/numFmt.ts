/**
 * Number formats: the built-in codes (ECMA-376 Part 1, §18.8.30), whether a
 * code formats dates or times, and serial dates.
 *
 * @internal
 */

/** Built-in number formats with a fixed code (ECMA-376 Part 1, §18.8.30). */
export const BUILTIN_FORMATS: ReadonlyMap<number, string> = new Map([
  [0, "General"],
  [1, "0"],
  [2, "0.00"],
  [3, "#,##0"],
  [4, "#,##0.00"],
  [9, "0%"],
  [10, "0.00%"],
  [11, "0.00E+00"],
  [12, "# ?/?"],
  [13, "# ??/??"],
  [14, "mm-dd-yy"],
  [15, "d-mmm-yy"],
  [16, "d-mmm"],
  [17, "mmm-yy"],
  [18, "h:mm AM/PM"],
  [19, "h:mm:ss AM/PM"],
  [20, "h:mm"],
  [21, "h:mm:ss"],
  [22, "m/d/yy h:mm"],
  [37, "#,##0 ;(#,##0)"],
  [38, "#,##0 ;[Red](#,##0)"],
  [39, "#,##0.00;(#,##0.00)"],
  [40, "#,##0.00;[Red](#,##0.00)"],
  [45, "mm:ss"],
  [46, "[h]:mm:ss"],
  [47, "mmss.0"],
  [48, "##0.0E+0"],
  [49, "@"]
])

/**
 * Built-in ids whose code depends on the locale: East Asian (27-36, 50-58)
 * and Thai (71-81) date and time formats. They format dates, but have no
 * fixed code.
 */
const LOCALE_DATE_IDS = new Set([
  27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  50, 51, 52, 53, 54, 55, 56, 57, 58,
  71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81
])

/**
 * Whether a format code shows a date or time: after dropping quoted text,
 * escaped characters (`\x`), padding and fill (`_x`, `*x`) and bracketed
 * parts (colors, conditions, locales), it still contains a date/time
 * letter (y, m, d, h, s, AM/PM). An elapsed-time bracket (`[h]`, `[mm]`,
 * `[ss]`) counts as time. `General` and number-only codes don't.
 */
export const isDateCode = (code: string): boolean => {
  let i = 0
  const n = code.length
  // progress: every branch moves i forward by at least 1
  while (i < n) {
    const c = code[i]!
    if (c === '"') {
      const close = code.indexOf('"', i + 1)
      i = close < 0 ? n : close + 1
    } else if (c === "\\" || c === "_" || c === "*") {
      i += 2
    } else if (c === "[") {
      const close = code.indexOf("]", i + 1)
      const inside = code.slice(i + 1, close < 0 ? n : close)
      if (/^(h+|m+|s+)$/i.test(inside)) return true
      i = close < 0 ? n : close + 1
    } else {
      const lower = c.toLowerCase()
      if (lower === "y" || lower === "m" || lower === "d" || lower === "h" || lower === "s") return true
      if ((lower === "a") && /^a(m\/pm|\/p)/i.test(code.slice(i, i + 5))) return true
      // "General" has no date letters; "E+" exponents don't either
      i++
    }
  }
  return false
}

export interface NumberFormat {
  /** The code, when known (custom, or a fixed built-in). */
  readonly code: string | undefined
  readonly date: boolean
}

const GENERAL: NumberFormat = { code: undefined, date: false }

/** The format for a numFmtId, given the workbook's custom codes. */
export const formatFor = (id: number, custom: ReadonlyMap<number, string>): NumberFormat => {
  const own = custom.get(id)
  if (own !== undefined) return { code: own, date: isDateCode(own) }
  if (id === 0) return GENERAL
  const builtin = BUILTIN_FORMATS.get(id)
  if (builtin !== undefined) return { code: builtin, date: isDateCode(builtin) }
  if (LOCALE_DATE_IDS.has(id)) return { code: undefined, date: true }
  return GENERAL
}

const DAY_MS = 86_400_000
const EPOCH_1900 = Date.UTC(1899, 11, 30)
const EPOCH_1900_EARLY = Date.UTC(1899, 11, 31)
const EPOCH_1904 = Date.UTC(1904, 0, 1)
/** Serial of 10000-01-01, the first day past Excel's date range, per system. */
const END_1900 = 2_958_466
const END_1904 = END_1900 - 1462

/**
 * A serial date as a `Date` (wall-clock time as UTC), rounded to the
 * millisecond; `undefined` outside Excel's range (negative, or past
 * 9999-12-31).
 *
 * 1900 system: serial 1 is 1900-01-01. Excel counts a 29 February 1900
 * that never existed (serial 60), so serials from 61 on are one day off a
 * plain day count, and serials below 60 use 1899-12-31 as day 0. Serial 60
 * itself (that day, with any time) has no `Date`: `undefined`, so the cell
 * stays a number.
 *
 * 1904 system: serial 0 is 1904-01-01.
 */
export const serialToDate = (serial: number, date1904: boolean): Date | undefined => {
  if (!Number.isFinite(serial) || serial < 0) return undefined
  if (date1904) return serial >= END_1904 ? undefined : new Date(EPOCH_1904 + Math.round(serial * DAY_MS))
  if (serial >= END_1900 || (serial >= 60 && serial < 61)) return undefined
  return new Date((serial < 60 ? EPOCH_1900_EARLY : EPOCH_1900) + Math.round(serial * DAY_MS))
}

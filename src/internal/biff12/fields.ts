/**
 * The string forms of BIFF12 record data, bounds-checked against the
 * record and against `maxStringLength`. Fixed-size fields are read through
 * a `FixedReader` (`binary/layout.ts`), once the record's length is checked.
 *
 * @internal
 */
import { utf16 } from "../biff/segments.ts"
import { ok, type StringRead, tooLongString, TRUNCATED } from "../binary/layout.ts"

/** A count of 0xFFFFFFFF: the null string of an XLNullableWideString. */
const NULL_COUNT = 0xFFFFFFFF

/** The u32 at `at`, if the data holds one there. */
const countAt = (data: Uint8Array, at: number): number | undefined =>
  at < 0 || at + 4 > data.length
    ? undefined
    : (data[at]! | (data[at + 1]! << 8) | (data[at + 2]! << 16) | (data[at + 3]! << 24)) >>> 0

/**
 * An XLWideString at `at`: a u32 count of UTF-16 code units, then the
 * units. `Truncated` if it runs past the data (whatever the limit), else
 * `TooLong` with the count if it is over `maxLength`; both are checked
 * before anything is decoded.
 */
export const readWideString = (data: Uint8Array, at: number, maxLength: number): StringRead => {
  const count = countAt(data, at)
  if (count === undefined) return TRUNCATED
  const end = at + 4 + 2 * count
  if (end > data.length) return TRUNCATED
  if (count > maxLength) return tooLongString(count)
  return ok(count === 0 ? "" : utf16.decode(data.subarray(at + 4, end)), end)
}

/**
 * An XLNullableWideString at `at`: an XLWideString, or a count of
 * 0xFFFFFFFF (no other data) for no string (`value` undefined).
 */
export const readNullableWideString = (data: Uint8Array, at: number, maxLength: number): StringRead<string | undefined> =>
  countAt(data, at) === NULL_COUNT ? ok(undefined, at + 4) : readWideString(data, at, maxLength)

/**
 * The text of a RichStr at `at` (a shared-string item, a rich cell): a
 * flags byte (formatting runs, phonetic data), then an XLWideString. The
 * runs and phonetic data after the text don't change it and aren't read.
 */
export const readRichStr = (data: Uint8Array, at: number, maxLength: number): StringRead =>
  at < 0 || at + 1 > data.length ? TRUNCATED : readWideString(data, at + 1, maxLength)

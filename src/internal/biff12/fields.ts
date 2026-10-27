/**
 * The string forms of BIFF12 record data, bounds-checked against the
 * record and against `maxStringLength`. Fixed-size fields are read with
 * the XLS reader's `Fields` (absolute offsets, `undefined` out of bounds).
 *
 * @internal
 */
import { utf16 } from "../biff/segments.ts"

export { Fields } from "../xls/fields.ts"

/** A string read at some offset, and the offset after it. */
export interface Read<A> {
  readonly value: A
  readonly end: number
}

/** A count of 0xFFFFFFFF: the null string of an XLNullableWideString. */
const NULL_COUNT = 0xFFFFFFFF

/**
 * An XLWideString at `at`: a u32 count of UTF-16 code units, then the
 * units. `undefined` if it runs past the data, `null` if the count is over
 * `maxLength` (checked before anything is decoded).
 */
export const readWideString = (data: Uint8Array, at: number, maxLength: number): Read<string> | undefined | null => {
  if (at < 0 || at + 4 > data.length) return undefined
  const count = (data[at]! | (data[at + 1]! << 8) | (data[at + 2]! << 16) | (data[at + 3]! << 24)) >>> 0
  if (count > maxLength) return null
  const end = at + 4 + 2 * count
  if (end > data.length) return undefined
  return { value: count === 0 ? "" : utf16.decode(data.subarray(at + 4, end)), end }
}

/**
 * An XLNullableWideString at `at`: an XLWideString, or a count of
 * 0xFFFFFFFF (no other data) for no string (`value` undefined).
 */
export const readNullableWideString = (
  data: Uint8Array,
  at: number,
  maxLength: number
): Read<string | undefined> | undefined | null => {
  if (at >= 0 && at + 4 <= data.length) {
    const count = (data[at]! | (data[at + 1]! << 8) | (data[at + 2]! << 16) | (data[at + 3]! << 24)) >>> 0
    if (count === NULL_COUNT) return { value: undefined, end: at + 4 }
  }
  return readWideString(data, at, maxLength)
}

/**
 * The text of a RichStr at `at` (a shared-string item, a rich cell): a
 * flags byte (formatting runs, phonetic data), then an XLWideString. The
 * runs and phonetic data after the text don't change it and aren't read.
 */
export const readRichStr = (data: Uint8Array, at: number, maxLength: number): Read<string> | undefined | null =>
  at < 0 || at + 1 > data.length ? undefined : readWideString(data, at + 1, maxLength)

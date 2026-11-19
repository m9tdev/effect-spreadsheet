/**
 * Reading data that spans a record and its `CONTINUE` records, as the
 * shared-string table (SST) does, through a cursor over them.
 *
 * Numbers may cross a segment boundary. Character data may too, but each
 * `CONTINUE` that continues a string starts with a fresh option byte saying
 * whether the rest of the characters are 8-bit or 16-bit (MS-XLS 2.5.293
 * XLUnicodeRichExtendedString, 2.1.4 records over 8224 bytes).
 *
 * Fixed fields are read after one check that enough bytes are left
 * (`has`), then without a check per field; character data checks its own
 * bounds and gives a typed result (`StringRead`). Every read moves the
 * cursor forward by what it consumes, so a loop over reads ends.
 *
 * @internal
 */

import { ok, type StringRead, tooLongString, TRUNCATED } from "../binary/layout.ts"
import { CONTINUE } from "./records.ts"

/** One UTF-16LE decoder for all 16-bit strings (decoding keeps no state between calls). */
export const utf16 = new TextDecoder("utf-16le")

/**
 * "Compressed" (8-bit) characters are the low bytes of UTF-16 code units:
 * byte 0x80 is U+0080, not the windows-1252 euro sign (a TextDecoder with
 * "latin1" or "iso-8859-1" would be windows-1252, per the Encoding spec).
 */
export const compressed = (bytes: Uint8Array): string => {
  let out = ""
  // progress: k increases by up to 8192 per step up to bytes.length
  for (let k = 0; k < bytes.length; k += 8192) {
    out += String.fromCharCode(...bytes.subarray(k, Math.min(bytes.length, k + 8192)))
  }
  return out
}

/**
 * The data of one logical record: a record's data, then the data of each
 * `CONTINUE` record right after it (MS-XLS 2.1.4: a record over 8 224
 * bytes goes on in `CONTINUE`s), read lazily. Each `CONTINUE`'s header is
 * looked at only when the reading gets to it (or `has` looks ahead to it),
 * so nothing is gathered or copied first: the cursor holds a few offsets,
 * whatever the number of `CONTINUE`s. Each record's data is a segment;
 * empty `CONTINUE`s continue nothing and are passed over.
 *
 * The record's `CONTINUE`s end at the first other record, or at the first
 * one that doesn't fit before `end` (the reader that read the record checks
 * them, and reports one that runs past its end). A record with no
 * `CONTINUE` after it, or `end` at its own end, is a cursor over one range.
 */
export class RecordCursor {
  private readonly bytes: Uint8Array
  /** Where the record's `CONTINUE`s must end. */
  private readonly end: number
  /** Start of the record's own data: a segment starting after it is a `CONTINUE`'s. */
  private readonly first: number
  /** The current segment, and where in it the next byte is read. */
  private segmentStart: number
  private segmentEnd: number
  private at: number
  /** The first header not looked at yet: every segment before it is counted in `known`. */
  private ahead: number
  /** Bytes from `at` up to `ahead`, over the segments between. */
  private known: number
  /** Bytes read so far, over all segments. */
  private read = 0

  /** The data of `record` (whose header is at `record.offset` in `bytes`) and of its `CONTINUE`s up to `end`. */
  constructor(bytes: Uint8Array, record: { readonly offset: number; readonly data: Uint8Array }, end: number) {
    this.bytes = bytes
    this.end = Math.min(end, bytes.length)
    this.first = record.offset + 4
    this.segmentStart = this.first
    this.segmentEnd = this.first + record.data.length
    this.at = this.first
    this.ahead = this.segmentEnd
    this.known = record.data.length
  }

  private u16At(at: number): number {
    return this.bytes[at]! | (this.bytes[at + 1]! << 8)
  }

  /**
   * Looks at the header at `ahead`: if it is a `CONTINUE` that fits before
   * `end`, counts its data in `known` and moves `ahead` past it; false if the
   * record's data ends there.
   */
  private scan(): boolean {
    const at = this.ahead
    if (at + 4 > this.end || this.u16At(at) !== CONTINUE) return false
    const next = at + 4 + this.u16At(at + 2)
    if (next > this.end) return false
    this.known += next - at - 4
    this.ahead = next
    return true
  }

  /** Moves to the next non-empty segment when the current one is used up; false at the end of the record's data. */
  private settle(): boolean {
    // progress: each pass moves to the next CONTINUE (segmentEnd grows by at least 4), up to `end`
    while (this.at >= this.segmentEnd) {
      const header = this.segmentEnd
      // a header before `ahead` was looked at already, and is a CONTINUE that fits
      if (header === this.ahead && !this.scan()) return false
      this.segmentStart = header + 4
      this.segmentEnd = this.segmentStart + this.u16At(header + 2)
      this.at = this.segmentStart
    }
    return true
  }

  /** Whether the cursor sits at the start of a `CONTINUE`'s data (a segment boundary). */
  private atBoundary(): boolean {
    return this.at === this.segmentStart && this.segmentStart !== this.first
  }

  /** Bytes read so far, over all segments: the offset of the cursor in their concatenation. */
  consumed(): number {
    return this.read
  }

  /**
   * Whether at least `n` bytes are left: the one check before reading `n`
   * bytes of fixed fields with `u8`, `u16` and `u32`. It looks at the
   * headers ahead only as far as it needs, each one once.
   */
  has(n: number): boolean {
    // progress: each scan moves `ahead` past one more CONTINUE, up to `end`
    while (this.known < n && this.scan()) {
      // counted by scan
    }
    return this.known >= n
  }

  /**
   * The next byte, across segments. Callers check `has` first; past the
   * end it reads 0 and moves nothing.
   */
  u8(): number {
    if (!this.settle()) return 0
    this.known--
    this.read++
    return this.bytes[this.at++]!
  }

  /** The next u16 (little-endian), across segments; checked by `has(2)` first. */
  u16(): number {
    // within the current segment (the common case): no settling per byte
    const at = this.at
    if (at + 2 <= this.segmentEnd) {
      this.at = at + 2
      this.known -= 2
      this.read += 2
      return this.u16At(at)
    }
    return this.u8() | (this.u8() << 8)
  }

  /** The next u32 (little-endian), across segments; checked by `has(4)` first. */
  u32(): number {
    return (this.u16() | (this.u16() << 16)) >>> 0
  }

  /** Skips `n` bytes across segments; false, skipping nothing, if fewer are left. */
  skip(n: number): boolean {
    if (!this.has(n)) return false
    let left = n
    // progress: each iteration skips at least 1 byte (n bytes are left, so
    // settle finds a non-empty segment)
    while (left > 0 && this.settle()) {
      const step = Math.min(left, this.segmentEnd - this.at)
      this.at += step
      this.known -= step
      this.read += step
      left -= step
    }
    return true
  }

  /**
   * Reads `count` characters, 16-bit when `highByte`, else 8-bit (Latin-1);
   * at each segment boundary inside the string, a new option byte gives
   * the width of the rest. `Truncated` if the segments end first, or end
   * a segment in the middle of a 16-bit character; `end` is `consumed()`.
   */
  chars(count: number, highByte: boolean): StringRead {
    let wide = highByte
    let left = count
    const parts: Array<string> = []
    // progress: each iteration consumes at least one character, or
    // boundary option bytes followed by at least one character, or fails
    while (left > 0) {
      if (!this.settle()) return TRUNCATED
      // characters continued in a new segment (also right after the
      // string's header ended a segment) start with an option byte; a
      // segment may hold only that byte, so the next one starts with its
      // own. progress: each pass reads one byte (settle found one)
      while (this.atBoundary()) {
        wide = (this.u8() & 0x01) !== 0
        if (!this.settle()) return TRUNCATED
      }
      const width = wide ? 2 : 1
      const fit = Math.floor((this.segmentEnd - this.at) / width)
      if (fit === 0) return TRUNCATED // half a 16-bit character at the end of a segment
      const take = Math.min(left, fit)
      const slice = this.bytes.subarray(this.at, this.at + take * width)
      parts.push(wide ? utf16.decode(slice) : compressed(slice))
      this.at += take * width
      this.known -= take * width
      this.read += take * width
      left -= take
    }
    return ok(parts.join(""), this.read)
  }
}

/**
 * Reads an XLUnicodeRichExtendedString (an SST entry): character count,
 * option flags, optional rich-text run count and phonetic block size, the
 * characters, then the run and phonetic data (skipped). `Truncated` if the
 * record's data ends early; `TooLong` with the character count when it is over
 * `maxLength` (found before any character is read). `end` is the reader's
 * `consumed()` after the string.
 *
 * After `TooLong` the reader sits before that string's characters, so the
 * caller must stop reading the SST there (fail with `LimitExceeded`): the
 * next read would start in the middle of a string.
 */
export const readRichString = (reader: RecordCursor, maxLength: number): StringRead => {
  // cch u16, flags u8, then cRun u16 and cbExtRst u32 when the flags say so
  if (!reader.has(3)) return TRUNCATED
  const count = reader.u16()
  const flags = reader.u8()
  const rich = (flags & 0x08) !== 0
  const ext = (flags & 0x04) !== 0
  if (!reader.has((rich ? 2 : 0) + (ext ? 4 : 0))) return TRUNCATED
  const runs = rich ? reader.u16() : 0
  const extSize = ext ? reader.u32() : 0
  if (count > maxLength) return tooLongString(count)
  const text = reader.chars(count, (flags & 0x01) !== 0)
  if (text._tag !== "Ok") return text
  if (!reader.skip(runs * 4) || !reader.skip(extSize)) return TRUNCATED
  return ok(text.value, reader.consumed())
}

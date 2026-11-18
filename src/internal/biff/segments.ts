/**
 * Reading data that spans a record and its `CONTINUE` records, as the
 * shared-string table (SST) does.
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
 * Segments without an object per segment: `count` byte ranges of `bytes`,
 * the k-th from `bounds[2k]` to `bounds[2k + 1]` (exclusive). Storage is 8
 * bytes per segment, whatever their number.
 */
export interface Segments {
  readonly bytes: Uint8Array
  readonly bounds: Uint32Array
  readonly count: number
}

/** Segments from separate arrays (copied into one buffer); for tests and small inputs. */
export const segmentsOf = (parts: ReadonlyArray<Uint8Array>): Segments => {
  let total = 0
  // progress: one part per iteration
  for (const part of parts) total += part.length
  const bytes = new Uint8Array(total)
  const bounds = new Uint32Array(2 * parts.length)
  let at = 0
  parts.forEach((part, k) => {
    bytes.set(part, at)
    bounds[2 * k] = at
    at += part.length
    bounds[2 * k + 1] = at
  })
  return { bytes, bounds, count: parts.length }
}

export class SegmentReader {
  private readonly bytes: Uint8Array
  private readonly bounds: Uint32Array
  private readonly count: number
  private segment = 0
  /** Offset within the current segment. */
  private at = 0
  /** Bytes in all segments. */
  private readonly total: number
  /** Bytes not yet read, over all segments. */
  private left: number

  constructor(segments: Segments | ReadonlyArray<Uint8Array>) {
    const s = "bounds" in segments ? segments : segmentsOf(segments)
    this.bytes = s.bytes
    this.bounds = s.bounds
    this.count = s.count
    let n = 0
    // progress: k up to count
    for (let k = 0; k < s.count; k++) n += s.bounds[2 * k + 1]! - s.bounds[2 * k]!
    this.total = n
    this.left = n
  }

  /** Bytes left in all segments (O(1): kept up to date by every read). */
  remaining(): number {
    return this.left
  }

  /** Bytes read so far, over all segments: the offset of the cursor in their concatenation. */
  consumed(): number {
    return this.total - this.left
  }

  /**
   * Whether at least `n` bytes are left: the one check before reading `n`
   * bytes of fixed fields with `u8`, `u16` and `u32`.
   */
  has(n: number): boolean {
    return this.left >= n
  }

  private length(k: number): number {
    return this.bounds[2 * k + 1]! - this.bounds[2 * k]!
  }

  /** Moves to the next non-empty segment when the current one is used up; false at the end. */
  private settle(): boolean {
    // progress: segment increases up to count
    while (this.segment < this.count && this.at >= this.length(this.segment)) {
      this.segment++
      this.at = 0
    }
    return this.segment < this.count
  }

  /** Whether the cursor sits at the start of a segment other than the first (a CONTINUE boundary). */
  private atBoundary(): boolean {
    return this.at === 0 && this.segment > 0
  }

  /**
   * The next byte, across segments. Callers check `has` first; past the
   * end it reads 0 and moves nothing.
   */
  u8(): number {
    if (!this.settle()) return 0
    this.left--
    return this.bytes[this.bounds[2 * this.segment]! + this.at++]!
  }

  /** The next u16 (little-endian), across segments; checked by `has(2)` first. */
  u16(): number {
    // within the current segment (the common case): no settling per byte
    const k = this.segment
    if (k < this.count) {
      const at = this.bounds[2 * k]! + this.at
      if (at + 2 <= this.bounds[2 * k + 1]!) {
        this.at += 2
        this.left -= 2
        return this.bytes[at]! | (this.bytes[at + 1]! << 8)
      }
    }
    return this.u8() | (this.u8() << 8)
  }

  /** The next u32 (little-endian), across segments; checked by `has(4)` first. */
  u32(): number {
    return (this.u16() | (this.u16() << 16)) >>> 0
  }

  /** Skips `n` bytes across segments; false, skipping nothing, if fewer are left. */
  skip(n: number): boolean {
    if (this.left < n) return false
    let left = n
    // progress: each iteration skips at least 1 byte (n bytes are left, so
    // settle finds a non-empty segment)
    while (left > 0 && this.settle()) {
      const step = Math.min(left, this.length(this.segment) - this.at)
      this.at += step
      this.left -= step
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
      const base = this.bounds[2 * this.segment]!
      const width = wide ? 2 : 1
      const fit = Math.floor((this.length(this.segment) - this.at) / width)
      if (fit === 0) return TRUNCATED // half a 16-bit character at the end of a segment
      const take = Math.min(left, fit)
      const slice = this.bytes.subarray(base + this.at, base + this.at + take * width)
      parts.push(wide ? utf16.decode(slice) : compressed(slice))
      this.at += take * width
      this.left -= take * width
      left -= take
    }
    return ok(parts.join(""), this.consumed())
  }
}

/**
 * Reads an XLUnicodeRichExtendedString (an SST entry): character count,
 * option flags, optional rich-text run count and phonetic block size, the
 * characters, then the run and phonetic data (skipped). `Truncated` if the
 * segments end early; `TooLong` with the character count when it is over
 * `maxLength` (found before any character is read). `end` is the reader's
 * `consumed()` after the string.
 *
 * After `TooLong` the reader sits before that string's characters, so the
 * caller must stop reading the SST there (fail with `LimitExceeded`): the
 * next read would start in the middle of a string.
 */
export const readRichString = (reader: SegmentReader, maxLength: number): StringRead => {
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

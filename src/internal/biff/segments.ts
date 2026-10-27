/**
 * Reading data that spans a record and its `CONTINUE` records, as the
 * shared-string table (SST) does.
 *
 * Numbers may cross a segment boundary. Character data may too, but each
 * `CONTINUE` that continues a string starts with a fresh option byte saying
 * whether the rest of the characters are 8-bit or 16-bit (MS-XLS 2.5.293
 * XLUnicodeRichExtendedString, 2.1.4 records over 8224 bytes).
 *
 * Every read moves the cursor forward by what it consumes; a read past the
 * last segment fails (the reader stays failed), so a loop over reads ends.
 *
 * @internal
 */

const utf16 = new TextDecoder("utf-16le")

/**
 * "Compressed" (8-bit) characters are the low bytes of UTF-16 code units:
 * byte 0x80 is U+0080, not the windows-1252 euro sign (a TextDecoder with
 * "latin1" or "iso-8859-1" would be windows-1252, per the Encoding spec).
 */
const compressed = (bytes: Uint8Array): string => {
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
  /** Bytes not yet read, over all segments. */
  private left: number
  /** Set when a read ran past the end; every later read fails too. */
  failed = false

  constructor(segments: Segments | ReadonlyArray<Uint8Array>) {
    const s = "bounds" in segments ? segments : segmentsOf(segments)
    this.bytes = s.bytes
    this.bounds = s.bounds
    this.count = s.count
    let n = 0
    // progress: k up to count
    for (let k = 0; k < s.count; k++) n += s.bounds[2 * k + 1]! - s.bounds[2 * k]!
    this.left = n
  }

  /** Bytes left in all segments (O(1): kept up to date by every read). */
  remaining(): number {
    return this.left
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

  u8(): number | undefined {
    if (this.failed || !this.settle()) {
      this.failed = true
      return undefined
    }
    this.left--
    return this.bytes[this.bounds[2 * this.segment]! + this.at++]
  }

  u16(): number | undefined {
    const a = this.u8()
    const b = this.u8()
    return a === undefined || b === undefined ? undefined : a | (b << 8)
  }

  u32(): number | undefined {
    const a = this.u16()
    const b = this.u16()
    return a === undefined || b === undefined ? undefined : (a | (b << 16)) >>> 0
  }

  /** Skips `n` bytes across segments; fails if fewer are left. */
  skip(n: number): boolean {
    let left = n
    // progress: each iteration skips at least 1 byte or ends
    while (left > 0) {
      if (this.failed || !this.settle()) {
        this.failed = true
        return false
      }
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
   * the width of the rest.
   */
  chars(count: number, highByte: boolean): string | undefined {
    let wide = highByte
    let left = count
    const parts: Array<string> = []
    // progress: each iteration consumes at least one character, or a
    // boundary option byte followed by at least one character, or fails
    while (left > 0) {
      if (this.failed || !this.settle()) {
        this.failed = true
        return undefined
      }
      // characters continued in a new segment (also right after the
      // string's header ended a segment) start with an option byte
      if (this.atBoundary()) {
        const option = this.u8()
        if (option === undefined || !this.settle()) {
          this.failed = true
          return undefined
        }
        wide = (option & 0x01) !== 0
      }
      const base = this.bounds[2 * this.segment]!
      const width = wide ? 2 : 1
      const fit = Math.floor((this.length(this.segment) - this.at) / width)
      if (fit === 0) {
        this.failed = true // half a 16-bit character at the end of a segment
        return undefined
      }
      const take = Math.min(left, fit)
      const slice = this.bytes.subarray(base + this.at, base + this.at + take * width)
      parts.push(wide ? utf16.decode(slice) : compressed(slice))
      this.at += take * width
      this.left -= take * width
      left -= take
    }
    return parts.join("")
  }
}

/**
 * Reads an XLUnicodeRichExtendedString (an SST entry): character count,
 * option flags, optional rich-text run count and phonetic block size, the
 * characters, then the run and phonetic data (skipped). `undefined` if the
 * segments end early; `null` when the character count is over `maxLength`.
 *
 * After `null` the reader sits before that string's characters, so the
 * caller must stop reading the SST there (fail with `LimitExceeded`): the
 * next read would start in the middle of a string.
 */
export const readRichString = (reader: SegmentReader, maxLength: number): string | undefined | null => {
  const count = reader.u16()
  const flags = reader.u8()
  if (count === undefined || flags === undefined) return undefined
  const rich = (flags & 0x08) !== 0
  const ext = (flags & 0x04) !== 0
  const runs = rich ? reader.u16() : 0
  const extSize = ext ? reader.u32() : 0
  if (runs === undefined || extSize === undefined) return undefined
  if (count > maxLength) return null
  const text = reader.chars(count, (flags & 0x01) !== 0)
  if (text === undefined) return undefined
  if (!reader.skip(runs * 4) || !reader.skip(extSize)) return undefined
  return text
}

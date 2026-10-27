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

const latin1 = new TextDecoder("windows-1252")
const utf16 = new TextDecoder("utf-16le")

export class SegmentReader {
  private readonly segments: ReadonlyArray<Uint8Array>
  private segment = 0
  private at = 0
  /** Set when a read ran past the end; every later read fails too. */
  failed = false

  constructor(segments: ReadonlyArray<Uint8Array>) {
    this.segments = segments
  }

  /** Bytes left in all segments. */
  remaining(): number {
    let n = 0
    // progress: k up to segments.length
    for (let k = this.segment; k < this.segments.length; k++) n += this.segments[k]!.length - (k === this.segment ? this.at : 0)
    return n
  }

  /** Moves to the next non-empty segment when the current one is used up; false at the end. */
  private settle(): boolean {
    // progress: segment increases up to segments.length
    while (this.segment < this.segments.length && this.at >= this.segments[this.segment]!.length) {
      this.segment++
      this.at = 0
    }
    return this.segment < this.segments.length
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
    return this.segments[this.segment]![this.at++]
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
      const step = Math.min(left, this.segments[this.segment]!.length - this.at)
      this.at += step
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
      const data = this.segments[this.segment]!
      const width = wide ? 2 : 1
      const fit = Math.floor((data.length - this.at) / width)
      if (fit === 0) {
        this.failed = true // half a 16-bit character at the end of a segment
        return undefined
      }
      const take = Math.min(left, fit)
      const slice = data.subarray(this.at, this.at + take * width)
      parts.push(wide ? utf16.decode(slice) : latin1.decode(slice))
      this.at += take * width
      left -= take
    }
    return parts.join("")
  }
}

/**
 * Reads an XLUnicodeRichExtendedString (an SST entry): character count,
 * option flags, optional rich-text run count and phonetic block size, the
 * characters, then the run and phonetic data (skipped). `undefined` if the
 * segments end early; `maxLength` caps the character count (`null` result
 * when it is exceeded).
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

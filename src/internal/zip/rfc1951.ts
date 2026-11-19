/**
 * Raw deflate decompression (RFC 1951), synchronous and pull-based, written
 * from the specification.
 *
 * The whole compressed buffer is given up front, so running out of input
 * always means the data is truncated, never "wait for more". Each `next`
 * call writes into the caller's buffer and stops when the buffer is full,
 * when it has consumed about `maxInput` compressed bytes, or at the end
 * of the stream. Bytes decoded before an error are always handed out (an
 * `Output` step) before the `Failed` step, so what a reader sees before an
 * error depends only on the input.
 *
 * Decoding keeps the last 32 KiB of output (the deflate window) in its own
 * buffer, since `out` may be a new buffer on each call.
 *
 * @internal
 */
import { SLICE_BYTES } from "../bytes.ts"
import type { InflateRaw } from "./inflate.ts"

export type InflateStep =
  /** `written` (at least 1) bytes were written at the start of `out`. */
  | { readonly _tag: "Output"; readonly written: number }
  /** The input budget was spent with no output (e.g. a run of empty blocks). */
  | { readonly _tag: "Working" }
  /** The stream ended; every byte has been handed out. */
  | { readonly _tag: "Done" }
  /** Invalid, truncated or trailing data at compressed byte `offset`; every byte decoded before it has been handed out. */
  | { readonly _tag: "Failed"; readonly reason: string; readonly offset: number }

export interface Inflate {
  /**
   * Works until it has filled `out` (which must not be empty), or consumed
   * `maxInput` compressed bytes (give or take the few bytes of one code or
   * block header), or the stream ends. After `Done` or `Failed`, returns
   * the same step again.
   */
  next(out: Uint8Array, maxInput: number): InflateStep
}

/** Starts decompressing `compressed`, a raw deflate stream. */
export const inflate = (compressed: Uint8Array): Inflate => new Inflater(compressed)

/** The deflate window: the farthest a match may reach back. */
const WINDOW = 32768

/** Index bits of the first-level lookup tables. */
const LIT_ROOT = 10
const DIST_ROOT = 8
const LIT_MASK = (1 << LIT_ROOT) - 1
const DIST_MASK = (1 << DIST_ROOT) - 1

// RFC 1951 3.2.5: base and extra bits of length symbols 257..285 and of distance symbols 0..29.
const LEN_BASE = Uint16Array.of(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
const LEN_EXTRA = Uint8Array.of(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
const DIST_BASE = Uint16Array.of(
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577
)
const DIST_EXTRA = Uint8Array.of(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13)
// RFC 1951 3.2.7: the order code length code lengths are written in.
const CL_ORDER = Uint8Array.of(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)

/*
 * Lookup tables. A code is read least-significant bit first, so a table is
 * indexed by the next `root` input bits, with each code's bits reversed.
 * An entry is one of:
 *   symbol << 8 | length        a code of `length` (1..15) bits in all
 *   offset << 8 | 0x80 | bits   codes longer than `root`: look the next
 *                               `bits` input bits up at `offset`
 *   0                           no code (an incomplete code set)
 */

/** How a code set built: complete, over-subscribed, incomplete, or with no codes at all. */
type Built = "complete" | "over" | "incomplete" | "single" | "empty"

/** Scratch space for `build`. */
const counts = new Uint16Array(16)
const nextCode = new Uint16Array(16)
const reversed = new Uint16Array(288)
const subBits = new Uint8Array(1 << LIT_ROOT)

/**
 * Builds the lookup table of the canonical code with these code lengths
 * (RFC 1951 3.2.2) into `holder.table`, growing it when needed. `single`:
 * incomplete, with exactly one code, of one bit.
 */
const build = (lengths: Uint8Array, root: number, holder: { table: Int32Array }): Built => {
  const n = lengths.length
  counts.fill(0)
  // progress: s up to n
  for (let s = 0; s < n; s++) counts[lengths[s]!]!++
  counts[0] = 0
  let max = 15
  // progress: max decreases to 0
  while (max > 0 && counts[max] === 0) max--
  const size = 1 << root
  if (max === 0) {
    if (holder.table.length < size) holder.table = new Int32Array(size)
    holder.table.fill(0, 0, size)
    return "empty"
  }
  let left = 1
  // progress: len up to 15
  for (let len = 1; len <= 15; len++) {
    left = (left << 1) - counts[len]!
    if (left < 0) return "over"
  }
  let code = 0
  // progress: len up to 15
  for (let len = 1; len <= 15; len++) {
    code = (code + counts[len - 1]!) << 1
    nextCode[len] = code
  }
  subBits.fill(0, 0, size)
  const mask = size - 1
  // progress: s up to n
  for (let s = 0; s < n; s++) {
    const len = lengths[s]!
    if (len === 0) continue
    let c = nextCode[len]!++
    let r = 0
    // progress: i up to len
    for (let i = 0; i < len; i++) {
      r = (r << 1) | (c & 1)
      c >>>= 1
    }
    reversed[s] = r
    if (len > root && len - root > subBits[r & mask]!) subBits[r & mask] = len - root
  }
  let total = size
  // progress: p up to size
  for (let p = 0; p < size; p++) if (subBits[p]! > 0) total += 1 << subBits[p]!
  if (holder.table.length < total) holder.table = new Int32Array(total)
  const table = holder.table
  table.fill(0, 0, total)
  let offset = size
  // progress: p up to size
  for (let p = 0; p < size; p++) {
    if (subBits[p]! > 0) {
      table[p] = (offset << 8) | 0x80 | subBits[p]!
      offset += 1 << subBits[p]!
    }
  }
  // progress: s up to n; each fill steps by at least 1
  for (let s = 0; s < n; s++) {
    const len = lengths[s]!
    if (len === 0) continue
    const r = reversed[s]!
    const entry = (s << 8) | len
    if (len <= root) {
      for (let i = r; i < size; i += 1 << len) table[i] = entry
    } else {
      const link = table[r & mask]!
      const start = link >>> 8
      const end = 1 << (link & 15)
      for (let i = r >>> root; i < end; i += 1 << (len - root)) table[start + i] = entry
    }
  }
  if (left === 0) return "complete"
  return max === 1 && counts[1] === 1 ? "single" : "incomplete"
}

/** The fixed codes of RFC 1951 3.2.6, built on first use. */
let fixed: { readonly lit: Int32Array; readonly dist: Int32Array } | undefined
const fixedTables = () => {
  if (fixed === undefined) {
    const lengths = new Uint8Array(288)
    lengths.fill(8, 0, 144)
    lengths.fill(9, 144, 256)
    lengths.fill(7, 256, 280)
    lengths.fill(8, 280, 288)
    const lit = { table: new Int32Array(0) }
    build(lengths, LIT_ROOT, lit)
    const dist = { table: new Int32Array(0) }
    build(new Uint8Array(32).fill(5), DIST_ROOT, dist)
    fixed = { lit: lit.table, dist: dist.table }
  }
  return fixed
}

/** Where decoding is: before a block header (or, after the last block, at the end), in a block, or finished. */
const Mode = { Header: 0, Stored: 1, Codes: 2, Done: 3, Failed: 4 } as const
type Mode = typeof Mode[keyof typeof Mode]

const DONE: InflateStep = { _tag: "Done" }
const WORKING: InflateStep = { _tag: "Working" }

class Inflater implements Inflate {
  private readonly input: Uint8Array
  /** Next input byte to load into `bits`; past the end, zero bytes are loaded (and counted). */
  private pos = 0
  /** Loaded bits not consumed yet, the next one lowest, and how many. */
  private bits = 0
  private count = 0
  private mode: Mode = Mode.Header
  private last = false
  /** Bytes left in the current stored block. */
  private storedLeft = 0
  private lit: Int32Array = new Int32Array(0)
  private dist: Int32Array = new Int32Array(0)
  private readonly litHolder = { table: new Int32Array(0) }
  private readonly distHolder = { table: new Int32Array(0) }
  private readonly lengths = new Uint8Array(286 + 30)
  /** A match that didn't fit in the last `out`: bytes still to copy, and its distance. */
  private copyLeft = 0
  private copyDist = 0
  /** The last `windowLength` (at most 32 KiB) bytes of output handed out before this call's `out`. */
  private readonly window = new Uint8Array(WINDOW)
  private windowLength = 0
  /** Bytes handed out so far, before this call's `out`. */
  private total = 0
  private failure: InflateStep = DONE

  constructor(input: Uint8Array) {
    this.input = input
  }

  next(out: Uint8Array, maxInput: number): InflateStep {
    if (this.mode === Mode.Done) return DONE
    if (this.mode === Mode.Failed) return this.failure
    if (out.length === 0) throw new RangeError("inflate needs a non-empty output buffer")
    const outLength = out.length
    // Budget in loaded bytes; at least one byte, so every call makes progress.
    const limit = this.pos + Math.max(1, maxInput)
    let o = 0
    // progress: each iteration writes output (o grows), consumes input
    // (pos grows: a block header, stored bytes, codes), or ends the loop
    // (out full, budget spent, end, error)
    while (o < outLength) {
      if (this.copyLeft > 0) {
        o = this.copy(out, o, this.copyLeft, this.copyDist)
        continue
      }
      const mode = this.mode
      if (mode === Mode.Codes) {
        if (this.pos >= limit) break
        o = this.codes(out, o, limit)
      } else if (mode === Mode.Stored) {
        if (this.storedLeft === 0) {
          this.mode = Mode.Header
          continue
        }
        if (this.pos >= limit) break
        o = this.stored(out, o, limit)
      } else if (mode === Mode.Header) {
        if (this.last) {
          this.end()
          break
        }
        if (this.pos >= limit) break
        this.header()
      } else break
    }
    if (o > 0) {
      this.keep(out, o)
      return { _tag: "Output", written: o }
    }
    // the methods called above may have ended the stream
    const after = this.mode as Mode
    if (after === Mode.Done) return DONE
    if (after === Mode.Failed) return this.failure
    return WORKING
  }

  /** Compressed bits consumed so far (zero bits loaded past the end included). */
  private consumed(): number {
    return this.pos * 8 - this.count
  }

  /** Fails, at the offset of the bit reached; as truncated if the data ran out before it. */
  private fail(reason: string, bit = this.consumed()): void {
    const truncated = this.consumed() > this.input.length * 8
    this.mode = Mode.Failed
    this.failure = {
      _tag: "Failed",
      reason: truncated ? "unexpected end of data" : reason,
      offset: truncated ? this.input.length : bit >>> 3
    }
  }

  /** Loads bytes until `bits` holds at least `n` (at most 25) bits. */
  private need(n: number): void {
    // progress: count grows by 8 each time, up to n
    while (this.count < n) {
      const at = this.pos++
      this.bits |= (at < this.input.length ? this.input[at]! : 0) << this.count
      this.count += 8
    }
  }

  /** The next `n` (at most 25) bits, consumed. */
  private take(n: number): number {
    this.need(n)
    const value = this.bits & ((1 << n) - 1)
    this.bits >>>= n
    this.count -= n
    return value
  }

  /** After the last block: the rest of the last byte is padding, and nothing may follow. */
  private end(): void {
    const consumed = this.consumed()
    if (consumed > this.input.length * 8) return this.fail("unexpected end of data")
    const used = (consumed + 7) >>> 3
    if (used < this.input.length) return this.fail("data after the end of the deflate stream", used * 8)
    this.mode = Mode.Done
  }

  /** Reads a block header and gets ready for its data. */
  private header(): void {
    const start = this.consumed()
    this.last = this.take(1) === 1
    const type = this.take(2)
    if (type === 0) {
      // stored: skip to a byte boundary, then LEN and its complement
      this.take(this.count & 7)
      const length = this.take(16)
      const complement = this.take(16)
      if (this.consumed() > this.input.length * 8) return this.fail("unexpected end of data")
      if ((length ^ 0xFFFF) !== complement) return this.fail("stored block length does not match its complement", start)
      // hand the whole bytes still loaded back to the input
      this.pos -= this.count >>> 3
      this.bits = 0
      this.count = 0
      this.storedLeft = length
      this.mode = Mode.Stored
    } else if (type === 1) {
      const tables = fixedTables()
      this.lit = tables.lit
      this.dist = tables.dist
      this.mode = Mode.Codes
    } else if (type === 2) {
      this.dynamic(start)
    } else {
      this.fail("invalid block type", start)
    }
  }

  /** Reads the code lengths of a dynamic block (RFC 1951 3.2.7) and builds its tables. */
  private dynamic(start: number): void {
    const litCount = this.take(5) + 257
    const distCount = this.take(5) + 1
    const clCount = this.take(4) + 4
    if (litCount > 286 || distCount > 30) return this.fail("too many length or distance codes", start)
    const clLengths = new Uint8Array(19)
    // progress: i up to clCount
    for (let i = 0; i < clCount; i++) clLengths[CL_ORDER[i]!] = this.take(3)
    const clHolder = { table: new Int32Array(1 << 7) }
    const clBuilt = build(clLengths, 7, clHolder)
    if (clBuilt !== "complete" && clBuilt !== "empty") return this.fail("invalid code length codes", start)
    const cl = clHolder.table
    const lengths = this.lengths
    const all = litCount + distCount
    let i = 0
    // progress: each code sets at least one length (i grows), or fails
    while (i < all) {
      this.need(7)
      const entry = cl[this.bits & 0x7F]!
      const len = entry & 15
      if (len === 0) return this.fail("invalid code length code")
      this.bits >>>= len
      this.count -= len
      const symbol = entry >>> 8
      if (symbol < 16) {
        lengths[i++] = symbol
        continue
      }
      let value = 0
      let repeat: number
      if (symbol === 16) {
        if (i === 0) return this.fail("length repeat with no previous length")
        value = lengths[i - 1]!
        repeat = 3 + this.take(2)
      } else if (symbol === 17) {
        repeat = 3 + this.take(3)
      } else {
        repeat = 11 + this.take(7)
      }
      if (i + repeat > all) return this.fail("code lengths repeat past their count")
      lengths.fill(value, i, i + repeat)
      i += repeat
    }
    if (this.consumed() > this.input.length * 8) return this.fail("unexpected end of data")
    if (lengths[256] === 0) return this.fail("no end-of-block code", start)
    const litBuilt = build(lengths.subarray(0, litCount), LIT_ROOT, this.litHolder)
    if (litBuilt === "over" || litBuilt === "incomplete") return this.fail("invalid literal/length code lengths", start)
    const distBuilt = build(lengths.subarray(litCount, all), DIST_ROOT, this.distHolder)
    if (distBuilt === "over" || distBuilt === "incomplete") return this.fail("invalid distance code lengths", start)
    this.lit = this.litHolder.table
    this.dist = this.distHolder.table
    this.mode = Mode.Codes
  }

  /** Copies stored bytes into `out` from `o`; the new `o`. */
  private stored(out: Uint8Array, o: number, limit: number): number {
    const input = this.input
    const n = Math.min(this.storedLeft, out.length - o, input.length - this.pos, Math.max(1, limit - this.pos))
    if (n <= 0) {
      this.fail("unexpected end of data")
      return o
    }
    out.set(input.subarray(this.pos, this.pos + n), o)
    this.pos += n
    this.storedLeft -= n
    return o + n
  }

  /**
   * Copies a match of `length` bytes from `distance` back into `out` from
   * `o`, as far as `out` holds; what doesn't fit is kept for the next call.
   * The new `o`.
   */
  private copy(out: Uint8Array, o: number, length: number, distance: number): number {
    const outLength = out.length
    let left = length
    if (distance > o) {
      // the first bytes come from the window, before `out`
      let from = this.windowLength - (distance - o)
      const n = Math.min(left, distance - o, outLength - o)
      const window = this.window
      // progress: o grows up to its bound
      for (const stop = o + n; o < stop;) out[o++] = window[from++]!
      left -= n
    }
    const n = Math.min(left, outLength - o)
    let from = o - distance
    // progress: o grows up to its bound
    for (const stop = o + n; o < stop;) out[o++] = out[from++]!
    left -= n
    this.copyLeft = left
    this.copyDist = distance
    return o
  }

  /** Decodes the codes of a compressed block into `out` from `o`, until `out` is full, the budget is spent, the block ends or an error. The new `o`. */
  private codes(out: Uint8Array, o: number, limit: number): number {
    const input = this.input
    const inLength = input.length
    const lit = this.lit
    const dist = this.dist
    const outLength = out.length
    let bits = this.bits
    let count = this.count
    let pos = this.pos
    // progress: each iteration consumes a code (at least one bit) and
    // writes output, ends the block or fails; stops when `out` is full or
    // the budget is spent
    while (o < outLength && pos < limit) {
      if (count < 16) {
        bits |= (pos + 1 < inLength ? input[pos]! | (input[pos + 1]! << 8) : (pos < inLength ? input[pos]! : 0)) << count
        pos += 2
        count += 16
      }
      let entry = lit[bits & LIT_MASK]!
      if ((entry & 0x80) !== 0) entry = lit[(entry >>> 8) + ((bits >>> LIT_ROOT) & ((1 << (entry & 15)) - 1))]!
      const len = entry & 15
      if (len === 0) {
        this.save(bits, count, pos)
        this.fail("invalid literal/length code")
        return o
      }
      bits >>>= len
      count -= len
      const symbol = entry >>> 8
      if (symbol < 256) {
        if (pos > inLength && (pos - inLength) * 8 > count) break // truncated: reported below
        out[o++] = symbol
        continue
      }
      if (symbol === 256) {
        this.save(bits, count, pos)
        if (this.consumed() > inLength * 8) {
          this.fail("unexpected end of data")
        } else {
          this.mode = Mode.Header
        }
        return o
      }
      const index = symbol - 257
      if (index >= 29) {
        this.save(bits, count, pos)
        this.fail("invalid literal/length code", this.consumed() - len)
        return o
      }
      // up to 5 extra bits, then a distance code of up to 15 bits: 20 bits
      // progress: count grows by 8 each time, up to 25 (32 bits at most)
      while (count <= 24) {
        bits |= (pos < inLength ? input[pos]! : 0) << count
        pos++
        count += 8
      }
      const lenExtra = LEN_EXTRA[index]!
      const length = LEN_BASE[index]! + (bits & ((1 << lenExtra) - 1))
      bits >>>= lenExtra
      count -= lenExtra
      let dEntry = dist[bits & DIST_MASK]!
      if ((dEntry & 0x80) !== 0) dEntry = dist[(dEntry >>> 8) + ((bits >>> DIST_ROOT) & ((1 << (dEntry & 15)) - 1))]!
      const dLen = dEntry & 15
      if (dLen === 0) {
        this.save(bits, count, pos)
        this.fail("invalid distance code")
        return o
      }
      bits >>>= dLen
      count -= dLen
      const dSymbol = dEntry >>> 8
      if (dSymbol >= 30) {
        this.save(bits, count, pos)
        this.fail("invalid distance code", this.consumed() - dLen)
        return o
      }
      if (count < 16) {
        bits |= ((pos < inLength ? input[pos]! : 0) | ((pos + 1 < inLength ? input[pos + 1]! : 0) << 8)) << count
        pos += 2
        count += 16
      }
      const dExtra = DIST_EXTRA[dSymbol]!
      const distance = DIST_BASE[dSymbol]! + (bits & ((1 << dExtra) - 1))
      bits >>>= dExtra
      count -= dExtra
      if (pos > inLength && (pos - inLength) * 8 > count) break // truncated: reported below
      if (distance > this.total + o) {
        this.save(bits, count, pos)
        this.fail("distance too far back")
        return o
      }
      if (distance <= o && length <= outLength - o) {
        // the common case: all in `out`
        let from = o - distance
        if (length >= 32 && distance >= length) {
          // long and not overlapping itself: one native copy
          out.copyWithin(o, from, from + length)
          o += length
        } else if (length >= 32 && distance === 1) {
          // a run of one byte
          out.fill(out[from]!, o, o + length)
          o += length
        } else {
          // progress: o grows up to its bound
          for (const stop = o + length; o < stop;) out[o++] = out[from++]!
        }
      } else {
        o = this.copy(out, o, length, distance)
        if (this.copyLeft > 0) break
      }
    }
    this.save(bits, count, pos)
    if (pos > inLength && this.consumed() > inLength * 8) this.fail("unexpected end of data")
    return o
  }

  private save(bits: number, count: number, pos: number): void {
    this.bits = bits
    this.count = count
    this.pos = pos
  }

  /** Keeps the last 32 KiB of output, `out`'s first `n` bytes being the newest. */
  private keep(out: Uint8Array, n: number): void {
    const window = this.window
    if (n >= WINDOW) {
      window.set(out.subarray(n - WINDOW, n))
      this.windowLength = WINDOW
    } else {
      const kept = Math.min(this.windowLength, WINDOW - n)
      window.copyWithin(0, this.windowLength - kept, this.windowLength)
      window.set(out.subarray(0, n), kept)
      this.windowLength = kept + n
    }
    this.total += n
  }
}

/**
 * This inflater behind the archive's `InflateRaw` interface, so it can be
 * swapped in for the default one and compared with it. Each `next` decodes
 * up to `SLICE_BYTES` from up to `SLICE_BYTES` of input; a call that only
 * consumed input gives an empty chunk. Errors reject with the reason and
 * the compressed offset, after the bytes decoded before them were given.
 */
export const inflateRawSync: InflateRaw = (compressed) => {
  const decoder = inflate(compressed)
  let stopped = false
  return {
    next: () => {
      if (stopped) return Promise.resolve(undefined)
      const out = new Uint8Array(SLICE_BYTES)
      const step = decoder.next(out, SLICE_BYTES)
      switch (step._tag) {
        case "Output":
          // a small chunk is copied out, so it doesn't keep the whole buffer alive
          return Promise.resolve(step.written * 2 < out.length ? out.slice(0, step.written) : out.subarray(0, step.written))
        case "Working":
          return Promise.resolve(new Uint8Array(0))
        case "Done":
          return Promise.resolve(undefined)
        case "Failed":
          return Promise.reject(new Error(`${step.reason} at compressed byte ${step.offset}`))
      }
    },
    cancel: () => {
      stopped = true
    }
  }
}

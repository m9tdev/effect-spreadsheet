/**
 * The TypeScript inflate (RFC 1951) against node:zlib (in tests only) and the
 * platform DecompressionStream: streams of every block type and setting,
 * hand-built streams for the edge cases of the format, truncated, corrupt
 * and trailing data, fuzzing, and the archive reading through it.
 */
import { Effect, Exit, Stream } from "effect"
import fc from "fast-check"
import * as zlib from "node:zlib"
import { describe, expect, it } from "vitest"
import { open } from "../src/internal/zip/archive.ts"
import { inflateRawWeb } from "../src/internal/zip/inflate.ts"
import { inflate, inflateRawSync, type InflateStep } from "../src/internal/zip/rfc1951.ts"
import { defaults as limits } from "../src/SpreadsheetLimits.ts"
import { buildZip, concat, testCrc32 } from "./helpers/zip.ts"

interface Run {
  readonly bytes: Buffer
  /** The last step: Done or Failed. */
  readonly end: InflateStep
  readonly steps: ReadonlyArray<InflateStep>
}

/** Runs the inflate to its end with outputs of `outSize` and input budgets of `maxInput`. */
const run = (data: Uint8Array, outSize = 65536, maxInput = 65536): Run => {
  const decoder = inflate(data)
  const parts: Array<Uint8Array> = []
  const steps: Array<InflateStep> = []
  // progress: each step writes, consumes input or ends (bounded for safety)
  for (let i = 0; i < 10_000_000; i++) {
    const out = new Uint8Array(outSize)
    const step = decoder.next(out, maxInput)
    steps.push(step)
    if (step._tag === "Output") parts.push(out.slice(0, step.written))
    if (step._tag === "Done" || step._tag === "Failed") return { bytes: Buffer.from(concat(parts)), end: step, steps }
  }
  throw new Error("inflate did not end")
}

/** zlib's result: its bytes (a truncated stream gives what it decoded, with no error) or its failure. */
const reference = (data: Uint8Array): { readonly ok: true; readonly bytes: Buffer } | { readonly ok: false } => {
  try {
    return { ok: true, bytes: zlib.inflateRawSync(data, { finishFlush: zlib.constants.Z_SYNC_FLUSH }) }
  } catch {
    return { ok: false }
  }
}

/** The platform inflater's result, and the bytes it gave before an error. */
const platform = async (data: Uint8Array): Promise<{ readonly ok: boolean; readonly bytes: Buffer }> => {
  const inflater = inflateRawWeb(data)
  const parts: Array<Uint8Array> = []
  try {
    // progress: each call returns a chunk or ends
    for (;;) {
      const chunk = await inflater.next()
      if (chunk === undefined) return { ok: true, bytes: Buffer.from(concat(parts)) }
      parts.push(chunk)
    }
  } catch {
    return { ok: false, bytes: Buffer.from(concat(parts)) }
  } finally {
    inflater.cancel()
  }
}

/**
 * Agrees with zlib: the same bytes when zlib succeeds; where zlib stops
 * early on a truncated stream, the same bytes then "unexpected end of data";
 * where zlib ignores trailing data, the same bytes then a failure; a
 * failure where zlib fails.
 */
const agreesWithZlib = (data: Uint8Array, result: Run) => {
  const z = reference(data)
  if (!z.ok) return result.end._tag === "Failed"
  if (!result.bytes.equals(z.bytes)) return false
  return result.end._tag === "Done" ||
    (result.end._tag === "Failed" && /^(unexpected end of data|data after the end)/.test(result.end.reason))
}

const pseudoRandom = (n: number, seed: number, mask = 0xFF) => {
  const out = new Uint8Array(n)
  let x = seed
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0
    out[i] = (x >>> 16) & mask
  }
  return out
}
const text = (n: number) => {
  const parts: Array<string> = []
  for (let i = 0; parts.length * 20 < n; i++) parts.push(`<c r="A${i}"><v>${i * 1.5}</v></c>`)
  return new TextEncoder().encode(parts.join("").slice(0, n))
}

/** Writes bits least-significant first, as deflate packs them; Huffman codes most-significant first. */
class Bits {
  private readonly bytes: Array<number> = []
  private bit = 0
  put(value: number, n: number) {
    for (let i = 0; i < n; i++) {
      if (this.bit === 0) this.bytes.push(0)
      this.bytes[this.bytes.length - 1]! |= ((value >>> i) & 1) << this.bit
      this.bit = (this.bit + 1) & 7
    }
    return this
  }
  code(code: number, length: number) {
    for (let i = length - 1; i >= 0; i--) this.put((code >>> i) & 1, 1)
    return this
  }
  align() {
    if (this.bit !== 0) this.bit = 0
    return this
  }
  raw(bytes: ArrayLike<number>) {
    this.align()
    for (let i = 0; i < bytes.length; i++) this.bytes.push(bytes[i]!)
    return this
  }
  done() {
    return Uint8Array.from(this.bytes)
  }
}

/** The canonical codes of these code lengths (RFC 1951 3.2.2). */
const canonical = (lengths: ReadonlyArray<number>): Array<number> => {
  const count = new Array<number>(16).fill(0)
  for (const l of lengths) if (l > 0) count[l]!++
  const next = new Array<number>(16).fill(0)
  let code = 0
  for (let l = 1; l <= 15; l++) {
    code = (code + count[l - 1]!) << 1
    next[l] = code
  }
  return lengths.map((l) => (l > 0 ? next[l]!++ : 0))
}

/** A fixed-code literal or length symbol (RFC 1951 3.2.6). */
const fixedLit = (w: Bits, symbol: number) =>
  symbol < 144
    ? w.code(0x30 + symbol, 8)
    : symbol < 256
    ? w.code(0x190 + symbol - 144, 9)
    : symbol < 280
    ? w.code(symbol - 256, 7)
    : w.code(0xC0 + symbol - 280, 8)

/** A stored block header with `length` (and `complement`, normally its ones' complement). */
const storedBlock = (w: Bits, last: boolean, bytes: ArrayLike<number>, complement = bytes.length ^ 0xFFFF) =>
  w.put(last ? 1 : 0, 1).put(0, 2).align().put(bytes.length, 16).put(complement, 16).raw(bytes)

/** Code length code lengths: a complete code over all 19 symbols (13 of 4 bits, 6 of 5). */
const CL_LENGTHS = Array.from({ length: 19 }, (_, s) => (s < 13 ? 4 : 5))
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]

/**
 * A dynamic block header: `lit` and `dist` code lengths, written with no
 * run-length codes, or with `codes` (code length symbols and their extra
 * bits) when given. The literal/length and distance codes, for writing data.
 */
const dynamicHeader = (
  w: Bits,
  last: boolean,
  lit: ReadonlyArray<number>,
  dist: ReadonlyArray<number>,
  codes?: ReadonlyArray<readonly [number, number, number]>
) => {
  const cl = canonical(CL_LENGTHS)
  w.put(last ? 1 : 0, 1).put(2, 2).put(lit.length - 257, 5).put(dist.length - 1, 5).put(19 - 4, 4)
  for (const s of CL_ORDER) w.put(CL_LENGTHS[s]!, 3)
  for (const [symbol, extra, extraBits] of codes ?? [...lit, ...dist].map((l) => [l, 0, 0] as const)) {
    w.code(cl[symbol]!, CL_LENGTHS[symbol]!)
    if (extraBits > 0) w.put(extra, extraBits)
  }
  return { lit: canonical(lit), dist: canonical(dist) }
}

describe("rfc1951 inflate: streams written by zlib", () => {
  const inputs: Array<readonly [string, Uint8Array]> = [
    ["empty", new Uint8Array(0)],
    ["one byte", Uint8Array.of(42)],
    ["one repeated byte, 300 KB", new Uint8Array(300_000).fill(7)],
    ["random bytes, 200 KB", pseudoRandom(200_000, 1)],
    ["low-entropy bytes, 200 KB", pseudoRandom(200_000, 2, 0x0F)],
    ["XML-like text, 500 KB", text(500_000)]
  ]
  const settings: Array<readonly [string, zlib.ZlibOptions]> = [
    ["level 0 (stored blocks)", { level: 0 }],
    ["level 1", { level: 1 }],
    ["level 6", { level: 6 }],
    ["level 9", { level: 9 }],
    ["fixed codes only", { strategy: zlib.constants.Z_FIXED }],
    ["Huffman only (no matches)", { strategy: zlib.constants.Z_HUFFMAN_ONLY }],
    ["run-length matches only", { strategy: zlib.constants.Z_RLE }],
    ["filtered", { strategy: zlib.constants.Z_FILTERED }],
    ["a 512-byte window", { windowBits: 9 }],
    ["small memory level (many blocks)", { memLevel: 1 }]
  ]
  for (const [label, data] of inputs) {
    it(`${label}: every setting, every step size`, () => {
      for (const [, options] of settings) {
        const compressed = zlib.deflateRawSync(data, options)
        for (const [outSize, maxInput] of [[65536, 65536], [1, 1], [4099, 37], [40000, 1_000_000]] as const) {
          const result = run(compressed, outSize, maxInput)
          expect(result.end._tag).toBe("Done")
          expect(result.bytes.equals(Buffer.from(data))).toBe(true)
        }
      }
    })
  }

  it("a stream written in many flushed pieces (empty stored blocks between them)", () => {
    const data = text(100_000)
    // zlib's sync flush ends each piece with an empty stored block
    const chunks: Array<Buffer> = []
    for (let i = 0; i < data.length; i += 1000) {
      chunks.push(zlib.deflateRawSync(Buffer.from(data.subarray(i, i + 1000)), { finishFlush: zlib.constants.Z_SYNC_FLUSH }))
    }
    // each piece is a complete set of blocks with no last-block flag; end with an empty last stored block
    const stream = Buffer.concat([...chunks, Buffer.from(storedBlock(new Bits(), true, []).done())])
    const result = run(stream, 3000, 100)
    expect(result.end._tag).toBe("Done")
    // pieces were compressed separately: each is decoded on its own
    expect(result.bytes.equals(Buffer.from(data))).toBe(true)
  })
})

describe("rfc1951 inflate: hand-built streams", () => {
  it("a long run of empty stored blocks is worked through within the input budget", () => {
    const w = new Bits()
    for (let i = 0; i < 2000; i++) storedBlock(w, false, [])
    storedBlock(w, true, [1, 2, 3])
    const stream = w.done()
    const result = run(stream, 65536, 64)
    expect(result.bytes.equals(Buffer.from([1, 2, 3]))).toBe(true)
    expect(result.end._tag).toBe("Done")
    // 5 bytes per empty block: one step per 64 to 69 input bytes (a block
    // started under the budget is finished), each without output
    const working = result.steps.filter((s) => s._tag === "Working").length
    expect(working).toBeGreaterThanOrEqual(Math.floor(stream.length / 69))
    expect(working).toBeLessThanOrEqual(Math.ceil(stream.length / 64))
    expect(agreesWithZlib(stream, result)).toBe(true)
  })

  it("an input budget below one block still makes progress", () => {
    const w = new Bits()
    for (let i = 0; i < 10; i++) storedBlock(w, false, [])
    storedBlock(w, true, [9])
    for (const budget of [0, -5, 1]) {
      const result = run(w.done(), 1, budget)
      expect(result.end._tag).toBe("Done")
      expect([...result.bytes]).toEqual([9])
    }
  })

  it("stored, fixed and dynamic blocks in one stream", () => {
    const w = new Bits()
    storedBlock(w, false, [65, 66])
    w.put(0, 1).put(1, 2) // fixed
    fixedLit(w, 67)
    fixedLit(w, 258) // length 4
    w.code(2, 5) // distance 3
    fixedLit(w, 256)
    const lit = new Array<number>(258).fill(0)
    lit[68] = 1
    lit[256] = 1
    const codes = dynamicHeader(w, true, lit, [0])
    w.code(codes.lit[68]!, 1).code(codes.lit[68]!, 1).code(codes.lit[256]!, 1)
    const stream = w.done()
    const result = run(stream)
    expect(result.bytes.toString("latin1")).toBe("ABCABCADD")
    expect(result.end._tag).toBe("Done")
    expect(agreesWithZlib(stream, result)).toBe(true)
  })

  it("a match reaching back exactly 32 KiB, split across outputs, and one byte farther fails", () => {
    const head = pseudoRandom(32768, 5)
    const make = (distanceExtra: number) => {
      const w = new Bits()
      storedBlock(w, false, head)
      w.put(1, 1).put(1, 2)
      fixedLit(w, 285) // length 258
      w.code(29, 5).put(distanceExtra, 13) // distance 24577 + extra
      fixedLit(w, 256)
      return w.done()
    }
    const ok = make(8191) // 32768
    for (const outSize of [65536, 1000, 1]) {
      const result = run(ok, outSize, 7)
      expect(result.end._tag).toBe("Done")
      expect(result.bytes.subarray(32768).equals(Buffer.from(head.subarray(0, 258)))).toBe(true)
    }
    expect(agreesWithZlib(ok, run(ok))).toBe(true)
    // the same stream with 1 byte less before it: 32768 is past the start
    const w = new Bits()
    storedBlock(w, false, head.subarray(1))
    w.put(1, 1).put(1, 2)
    fixedLit(w, 285)
    w.code(29, 5).put(8191, 13)
    fixedLit(w, 256)
    const far = w.done()
    const result = run(far)
    expect(result.end).toMatchObject({ _tag: "Failed", reason: "distance too far back" })
    // the stored bytes before the error are handed out
    expect(result.bytes.equals(Buffer.from(head.subarray(1)))).toBe(true)
    expect(reference(far).ok).toBe(false)
  })

  it("a long match that overlaps itself repeats its source, in one output or across outputs", () => {
    const w = new Bits().put(1, 1).put(1, 2)
    const head = Array.from({ length: 30 }, (_, i) => 97 + (i % 26))
    for (const c of head) fixedLit(w, c)
    fixedLit(w, 272).put(3, 2) // length 31 + 3
    w.code(9, 5).put(5, 3) // distance 25 + 5
    fixedLit(w, 256)
    const expected = Buffer.from([...head, ...head, ...head.slice(0, 4)])
    for (const outSize of [65536, 40, 1]) {
      const result = run(w.done(), outSize)
      expect(result.end._tag).toBe("Done")
      expect(result.bytes.equals(expected)).toBe(true)
    }
    expect(agreesWithZlib(w.done(), run(w.done()))).toBe(true)
  })

  it("a stream cut right after a final end-of-block code read from padding fails", () => {
    // a fixed block whose end-of-block code (7 zero bits) is entirely past the end
    const w = new Bits().put(1, 1).put(1, 2)
    fixedLit(w, 65)
    const stream = w.done() // 3 + 8 bits: the code would be read from padding
    const result = run(stream)
    expect(result.end).toEqual({ _tag: "Failed", reason: "unexpected end of data", offset: stream.length })
    expect(result.bytes.toString("latin1")).toBe("A")
  })

  it("a distance past the start of the output fails, after the bytes before it", () => {
    const w = new Bits().put(1, 1).put(1, 2)
    fixedLit(w, 65)
    fixedLit(w, 66)
    fixedLit(w, 257) // length 3
    w.code(2, 5) // distance 3, with 2 bytes out
    fixedLit(w, 256)
    const result = run(w.done())
    expect(result.end).toMatchObject({ _tag: "Failed", reason: "distance too far back" })
    expect(result.bytes.toString("latin1")).toBe("AB")
    expect(result.steps.map((s) => s._tag)).toEqual(["Output", "Failed"])
    expect(reference(w.done()).ok).toBe(false)
  })

  const fails: Array<readonly [string, () => Uint8Array, string, number]> = [
    ["block type 3", () => new Bits().put(1, 1).put(3, 2).done(), "invalid block type", 0],
    ["a stored length that doesn't match its complement", () => storedBlock(new Bits(), true, [1], 0).done(), "stored block length does not match its complement", 0],
    ["fixed literal/length code 286", () => fixedLit(new Bits().put(1, 1).put(1, 2), 286).done(), "invalid literal/length code", 0],
    ["fixed literal/length code 287", () => fixedLit(new Bits().put(1, 1).put(1, 2), 287).done(), "invalid literal/length code", 0],
    ["fixed distance code 30", () => {
      const w = new Bits().put(1, 1).put(1, 2)
      fixedLit(w, 65)
      fixedLit(w, 257)
      return w.code(30, 5).done()
    }, "invalid distance code", 2],
    ["more than 286 literal/length codes", () => new Bits().put(1, 1).put(2, 2).put(30, 5).put(0, 5).put(15, 4).done(), "too many length or distance codes", 0],
    ["more than 30 distance codes", () => new Bits().put(1, 1).put(2, 2).put(0, 5).put(30, 5).put(15, 4).done(), "too many length or distance codes", 0],
    ["a repeat of the previous length with none before it", () => {
      const w = new Bits()
      dynamicHeader(w, true, new Array(257).fill(8), [1], [[16, 0, 2]])
      return w.done()
    }, "length repeat with no previous length", 9], // header 74 bits, then the 5-bit repeat code
    ["a repeat past the number of lengths", () => {
      const w = new Bits()
      dynamicHeader(w, true, new Array(257).fill(8), [1], [[8, 0, 0], [18, 127, 7], [18, 127, 7], [18, 127, 7]])
      return w.done()
    }, "code lengths repeat past their count", 12], // 74 + 4 + 12 + 12 bits
    ["no end-of-block code", () => {
      const lit = new Array<number>(257).fill(0)
      lit[0] = 1
      lit[1] = 1
      const w = new Bits()
      dynamicHeader(w, true, lit, [1])
      return w.done()
    }, "no end-of-block code", 0],
    ["over-subscribed literal/length codes", () => {
      const lit = new Array<number>(257).fill(0)
      lit[0] = 1
      lit[1] = 1
      lit[256] = 1
      const w = new Bits()
      dynamicHeader(w, true, lit, [1])
      return w.done()
    }, "invalid literal/length code lengths", 0],
    ["incomplete literal/length codes", () => {
      const lit = new Array<number>(257).fill(0)
      lit[0] = 2
      lit[256] = 2
      const w = new Bits()
      dynamicHeader(w, true, lit, [1])
      return w.done()
    }, "invalid literal/length code lengths", 0],
    ["incomplete distance codes (two codes of 2 bits)", () => {
      const lit = new Array<number>(257).fill(0)
      lit[0] = 1
      lit[256] = 1
      const w = new Bits()
      dynamicHeader(w, true, lit, [2, 2])
      return w.done()
    }, "invalid distance code lengths", 0],
    ["incomplete code length codes", () => {
      const w = new Bits().put(1, 1).put(2, 2).put(0, 5).put(0, 5).put(0, 4)
      // 4 code length codes, of lengths 1, 2, 0, 0: incomplete
      return w.put(1, 3).put(2, 3).put(0, 3).put(0, 3).put(0, 32).done()
    }, "invalid code length codes", 0]
  ]
  for (const [label, make, reason, offset] of fails) {
    it(`fails on ${label}, at its offset, as zlib does`, () => {
      const stream = make()
      // some padding, so the data isn't simply cut short
      const padded = concat([stream, new Uint8Array(8)])
      const result = run(padded)
      expect(result.end).toEqual({ _tag: "Failed", reason, offset })
      expect(reference(padded).ok).toBe(false)
    })
  }

  it("a single one-bit distance code is allowed; its unused code fails", () => {
    const lit = new Array<number>(258).fill(0)
    lit[65] = 1
    lit[256] = 2
    lit[257] = 2
    const make = (distanceBit: number) => {
      const w = new Bits()
      const codes = dynamicHeader(w, true, lit, [1])
      w.code(codes.lit[65]!, 1).code(codes.lit[257]!, 2).put(distanceBit, 1).code(codes.lit[256]!, 2)
      return w.done()
    }
    const ok = run(make(0))
    expect([ok.bytes.toString("latin1"), ok.end._tag]).toEqual(["AAAA", "Done"])
    expect(agreesWithZlib(make(0), ok)).toBe(true)
    const bad = run(make(1))
    expect(bad.end).toMatchObject({ _tag: "Failed", reason: "invalid distance code" })
    expect(bad.bytes.toString("latin1")).toBe("A")
    expect(reference(make(1)).ok).toBe(false)
  })

  it("Done and Failed are given again on later calls", () => {
    const ok = inflate(zlib.deflateRawSync(Buffer.from("abc")))
    const out = new Uint8Array(10)
    expect(ok.next(out, 100)).toEqual({ _tag: "Output", written: 3 })
    expect(ok.next(out, 100)).toEqual({ _tag: "Done" })
    expect(ok.next(out, 100)).toEqual({ _tag: "Done" })
    const bad = inflate(Uint8Array.of(0xFF))
    const first = bad.next(out, 100)
    expect(first._tag).toBe("Failed")
    expect(bad.next(out, 100)).toEqual(first)
  })

  it("an empty output buffer is a programming error", () => {
    expect(() => inflate(Uint8Array.of(3, 0)).next(new Uint8Array(0), 10)).toThrow(RangeError)
  })
})

describe("rfc1951 inflate: truncated, corrupt and trailing data", () => {
  const data = text(70_000)
  const streams = [
    ["dynamic", zlib.deflateRawSync(data)],
    ["fixed", zlib.deflateRawSync(data, { strategy: zlib.constants.Z_FIXED })],
    ["stored", zlib.deflateRawSync(data.subarray(0, 3000), { level: 0 })]
  ] as const

  for (const [label, stream] of streams) {
    it(`${label}: every cut fails with "unexpected end of data" after exactly the bytes zlib decodes`, () => {
      const step = Math.max(1, Math.floor(stream.length / 400))
      for (let cut = 0; cut < stream.length; cut += step) {
        const part = stream.subarray(0, cut)
        const result = run(part, 4096, 1000)
        expect(result.end).toEqual({ _tag: "Failed", reason: "unexpected end of data", offset: cut })
        const z = reference(part)
        expect(z.ok && result.bytes.equals(z.bytes)).toBe(true)
      }
    })

    it(`${label}: data after the end fails after every byte, at the end of the stream`, () => {
      const result = run(concat([stream, Uint8Array.of(0)]))
      expect(result.end).toEqual({ _tag: "Failed", reason: "data after the end of the deflate stream", offset: stream.length })
      expect(result.bytes.equals(Buffer.from(label === "stored" ? data.subarray(0, 3000) : data))).toBe(true)
    })
  }

  it("the bytes before an error don't depend on the step sizes", () => {
    const stream = zlib.deflateRawSync(text(200_000))
    fc.assert(
      fc.property(fc.nat(stream.length * 8 - 1), fc.integer({ min: 1, max: 70_000 }), fc.integer({ min: 1, max: 70_000 }), (bit, outSize, maxInput) => {
        const flipped = stream.slice()
        flipped[bit >>> 3]! ^= 1 << (bit & 7)
        const a = run(flipped)
        const b = run(flipped, outSize, maxInput)
        expect(b.bytes.equals(a.bytes)).toBe(true)
        expect(b.end).toEqual(a.end)
      }),
      { numRuns: 60 }
    )
  })

  it("fuzz: random bytes end in Done or Failed, as zlib's result", () => {
    fc.assert(
      fc.property(fc.uint8Array({ maxLength: 2000 }), fc.integer({ min: 1, max: 5000 }), (bytes, outSize) => {
        const result = run(bytes, outSize, 100)
        expect(agreesWithZlib(bytes, result)).toBe(true)
      }),
      { numRuns: 3000 }
    )
  })

  it("fuzz: bit flips, cuts and inserted bytes in valid streams agree with zlib", () => {
    const valid = [
      zlib.deflateRawSync(text(20_000)),
      zlib.deflateRawSync(text(20_000), { strategy: zlib.constants.Z_FIXED }),
      zlib.deflateRawSync(pseudoRandom(5000, 3, 0x3F), { level: 9 }),
      zlib.deflateRawSync(text(5000), { level: 0 })
    ]
    fc.assert(
      fc.property(
        fc.nat(valid.length - 1),
        fc.array(fc.tuple(fc.nat(), fc.nat(7), fc.nat(255), fc.constantFrom("flip", "set", "insert", "cut")), { minLength: 1, maxLength: 4 }),
        (which, edits) => {
          let bytes = valid[which]!.slice() as Uint8Array
          for (const [at, bit, value, edit] of edits) {
            if (bytes.length === 0) break
            const i = at % bytes.length
            if (edit === "flip") bytes[i]! ^= 1 << bit
            else if (edit === "set") bytes[i] = value
            else if (edit === "insert") bytes = concat([bytes.subarray(0, i), Uint8Array.of(value), bytes.subarray(i)])
            else bytes = bytes.subarray(0, i)
          }
          expect(agreesWithZlib(bytes, run(bytes, 777, 333))).toBe(true)
        }
      ),
      { numRuns: 1500 }
    )
  })
})

describe("rfc1951 inflate: against DecompressionStream", () => {
  it("valid streams give the same bytes; broken ones fail in both, the platform's bytes a prefix of ours", async () => {
    const data = text(300_000)
    const valid = zlib.deflateRawSync(data)
    const cases: Array<Uint8Array> = [valid, zlib.deflateRawSync(data, { level: 0 }), zlib.deflateRawSync(data, { strategy: zlib.constants.Z_FIXED })]
    for (const cut of [1, 100, 5000, valid.length >>> 1, valid.length - 1]) cases.push(valid.subarray(0, cut))
    cases.push(concat([valid, Uint8Array.of(1, 2, 3, 4)]))
    for (let i = 0; i < 20; i++) {
      const flipped = valid.slice()
      flipped[(i * 7919) % valid.length]! ^= 1 << (i % 8)
      cases.push(flipped)
    }
    for (const bytes of cases) {
      const p = await platform(bytes)
      const t = run(bytes)
      expect(t.end._tag === "Done").toBe(p.ok)
      if (p.ok) expect(t.bytes.equals(p.bytes)).toBe(true)
      else expect(t.bytes.subarray(0, p.bytes.length).equals(p.bytes)).toBe(true)
    }
  })
})

describe("rfc1951 inflate behind the archive's inflate interface", () => {
  it("an archive read through it gives the same bytes as through the default", async () => {
    const big = text(400_000)
    const zip = await buildZip([
      { name: "a.xml", data: big },
      { name: "b.bin", data: pseudoRandom(100_000, 9) },
      { name: "empty.xml", data: new Uint8Array(0) }
    ])
    for (const name of ["a.xml", "b.bin", "empty.xml"]) {
      const read = async (inflater?: typeof inflateRawSync) => {
        const opened = inflater === undefined ? open(zip, limits) : open(zip, limits, inflater)
        if (opened._tag !== "Success") throw new Error("open failed")
        const archive = opened.success
        const chunks = await Effect.runPromise(Stream.runCollect(archive.read(archive.find(name)!, "test")))
        return Buffer.from(concat(Array.from(chunks)))
      }
      expect((await read(inflateRawSync)).equals(await read())).toBe(true)
    }
  })

  it("a damaged entry fails with the reason and its compressed offset", async () => {
    const data = text(10_000)
    const zip = await buildZip([{ name: "a.xml", data, method: "store", methodCode: 8, crc32: testCrc32(data) }])
    const opened = open(zip, limits, inflateRawSync)
    if (opened._tag !== "Success") throw new Error("open failed")
    const archive = opened.success
    const exit = await Effect.runPromiseExit(Stream.runCollect(archive.read(archive.find("a.xml")!, "test")))
    expect(Exit.isFailure(exit)).toBe(true)
    expect(JSON.stringify(exit)).toMatch(/invalid deflate data: [a-z /-]+ at compressed byte \d+/)
  })
})

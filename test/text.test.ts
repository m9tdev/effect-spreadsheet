import * as Result from "effect/Result"
import fc from "fast-check"
import { describe, expect, it } from "vitest"
import { bomEncoding, DETECT_BYTES, detectEncoding, firstInvalidUtf8, TextInput } from "../src/internal/text.ts"

/** Whether a fatal `TextDecoder` accepts `bytes` (in stream mode unless `complete`). */
const platformValid = (bytes: Uint8Array, complete: boolean): boolean => {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes, { stream: !complete })
    return true
  } catch {
    return false
  }
}

const lenient = (bytes: Uint8Array): string => new TextDecoder("utf-8").decode(bytes)

/**
 * Bytes near UTF-8's edges: ASCII, every lead byte class, continuation
 * bytes at the bounds of the restricted ranges, and invalid bytes.
 */
const edgeByte = fc.constantFrom(
  0x00, 0x41, 0x7F, 0x80, 0x8F, 0x90, 0x9F, 0xA0, 0xBF, 0xC0, 0xC1, 0xC2, 0xDF, 0xE0, 0xE1, 0xEC, 0xED, 0xEE,
  0xEF, 0xF0, 0xF1, 0xF3, 0xF4, 0xF5, 0xF8, 0xFE, 0xFF, 0xBB, 0xBD
)
const edgeBytes = fc.array(edgeByte, { maxLength: 24 }).map((a) => new Uint8Array(a))
/** Valid text with some bytes replaced: mostly valid, with an error somewhere. */
const nearlyValid = fc.tuple(fc.string({ unit: "grapheme", maxLength: 16 }), fc.array(fc.tuple(fc.nat(), edgeByte), { maxLength: 2 }))
  .map(([s, edits]) => {
    const bytes = new TextEncoder().encode(s)
    for (const [at, b] of edits) if (bytes.length > 0) bytes[at % bytes.length] = b
    return bytes
  })
const someBytes = fc.oneof(edgeBytes, nearlyValid)

/** `bytes` cut at `cuts` (taken modulo its length + 1, sorted). */
const split = (bytes: Uint8Array, cuts: ReadonlyArray<number>): Array<Uint8Array> => {
  const at = [...new Set(cuts.map((c) => c % (bytes.length + 1)))].sort((a, b) => a - b)
  const out: Array<Uint8Array> = []
  let from = 0
  for (const c of at) {
    out.push(bytes.subarray(from, c))
    from = c
  }
  out.push(bytes.subarray(from))
  return out
}

type Outcome = { readonly text: string } | { readonly invalid: number | undefined }

/** Runs `input` over `pieces`: the text, or the offset its first error had. */
const run = (input: TextInput<number | undefined>, pieces: ReadonlyArray<Uint8Array>): Outcome => {
  let text = ""
  for (const piece of pieces) {
    const out = input.push(piece)
    if (Result.isFailure(out)) return { invalid: out.failure }
    text += out.success
  }
  const rest = input.end()
  return Result.isFailure(rest) ? { invalid: rest.failure } : { text: text + rest.success }
}

const fatal = (encoding = "utf-8") =>
  new TextInput<number | undefined>({ detect: "bom-only", encoding, fatal: true, invalid: (offset) => offset })

const bytes = (...b: Array<number>) => new Uint8Array(b)

describe("text: firstInvalidUtf8", () => {
  it("finds the first byte of the first ill-formed sequence, at the edges of every range", () => {
    const cases: Array<[Array<number>, number]> = [
      [[0x41, 0x7F], -1],
      [[0x80], 0], // a lone continuation byte
      [[0x41, 0xBF], 1],
      [[0xC0, 0x80], 0], // overlong lead bytes
      [[0xC1, 0xBF], 0],
      [[0xC2, 0x80], -1],
      [[0xDF, 0xBF], -1],
      [[0xC2, 0x41], 0],
      [[0xC2, 0xC0], 0],
      [[0xE0, 0x9F, 0x80], 0], // overlong
      [[0xE0, 0xA0, 0x80], -1],
      [[0xE1, 0x80, 0x80], -1],
      [[0xED, 0x9F, 0xBF], -1],
      [[0xED, 0xA0, 0x80], 0], // a surrogate
      [[0xEF, 0xBF, 0xBD], -1], // U+FFFD itself
      [[0xE1, 0x80, 0x41], 0], // the third byte isn't a continuation
      [[0xF0, 0x8F, 0x80, 0x80], 0], // overlong
      [[0xF0, 0x90, 0x80, 0x80], -1],
      [[0xF4, 0x8F, 0xBF, 0xBF], -1],
      [[0xF4, 0x90, 0x80, 0x80], 0], // past U+10FFFF
      [[0xF1, 0x80, 0x80, 0x41], 0], // the fourth byte isn't a continuation
      [[0xF5, 0x80, 0x80, 0x80], 0],
      [[0xFF], 0],
      [[0x41, 0x42, 0xC3, 0xA9, 0xFE], 4]
    ]
    for (const [b, at] of cases) {
      expect(firstInvalidUtf8(new Uint8Array(b), true), JSON.stringify(b)).toBe(at)
      expect(firstInvalidUtf8(new Uint8Array(b), false), JSON.stringify(b)).toBe(at)
    }
  })

  it("a sequence cut at the end is ill-formed only when the input is complete, and only if it could go on", () => {
    for (const b of [[0xC3], [0xE2, 0x82], [0xF0, 0x9F, 0x98], [0x41, 0xF4, 0x8F]]) {
      expect(firstInvalidUtf8(new Uint8Array(b), false)).toBe(-1)
      expect(firstInvalidUtf8(new Uint8Array(b), true)).toBe(b.length - (b[0] === 0x41 ? 2 : b.length))
    }
    // already wrong before the end
    expect(firstInvalidUtf8(bytes(0xE0, 0x80), false)).toBe(0)
    expect(firstInvalidUtf8(bytes(0xF4, 0x90), false)).toBe(0)
  })

  it("agrees with a fatal TextDecoder, and the offset is where the lenient decoder's first U+FFFD comes from", () => {
    fc.assert(
      fc.property(someBytes, fc.boolean(), (b, complete) => {
        const at = firstInvalidUtf8(b, complete)
        expect(at === -1).toBe(platformValid(b, complete))
        if (at >= 0) {
          // everything before is whole, valid text, and the error starts right after it
          const before = b.subarray(0, at)
          expect(platformValid(before, true)).toBe(true)
          const prefix = lenient(before)
          const all = complete ? lenient(b) : new TextDecoder("utf-8").decode(b, { stream: true })
          expect(all.startsWith(prefix)).toBe(true)
          expect(all.charAt(prefix.length)).toBe("\uFFFD")
        }
      }),
      { numRuns: 3000 }
    )
  })
})

describe("text: encoding detection", () => {
  it("UTF-8 when valid, windows-1252 when not; a sequence cut at the end counts only when the input is complete", () => {
    expect(detectEncoding(new TextEncoder().encode("né,ü"), true)).toBe("utf-8")
    expect(detectEncoding(bytes(0x6E, 0xE9), true)).toBe("windows-1252")
    expect(detectEncoding(bytes(0x6E, 0xC3), true)).toBe("windows-1252")
    expect(detectEncoding(bytes(0x6E, 0xC3), false)).toBe("utf-8")
    expect(detectEncoding(bytes(0xEF, 0xBB, 0xBF, 0xE9), true)).toBe("utf-8")
    expect(bomEncoding(bytes(0xFE, 0xFF))).toBe("utf-16be")
  })
})

describe("text: TextInput", () => {
  it("fatal UTF-8: the same text, or the offset of the first invalid sequence, however the input is split", () => {
    fc.assert(
      fc.property(fc.boolean(), someBytes, fc.array(fc.nat(), { maxLength: 6 }), (bom, body, cuts) => {
        const input = bom ? new Uint8Array([0xEF, 0xBB, 0xBF, ...body]) : body
        // a UTF-16 byte-order mark picks UTF-16
        fc.pre(!(input.length >= 2 && (input[0] === 0xFF || input[0] === 0xFE) && input[0] + input[1]! === 0x1FD))
        const at = firstInvalidUtf8(input, true)
        const expected: Outcome = at >= 0 ? { invalid: at } : { text: lenient(input) }
        expect(run(fatal(), split(input, cuts))).toEqual(expected)
      }),
      { numRuns: 3000 }
    )
  })

  it("fatal UTF-8 accepts U+FFFD in the input, before and after a split", () => {
    const input = new TextEncoder().encode("a\uFFFDb\uFFFD")
    for (let cut = 0; cut <= input.length; cut++) {
      expect(run(fatal(), split(input, [cut]))).toEqual({ text: "a\uFFFDb\uFFFD" })
    }
    // ... but not the invalid byte after it
    expect(run(fatal(), [input, bytes(0x80)])).toEqual({ invalid: input.length })
  })

  it("fatal UTF-8: a sequence cut at the end of the input is an error at its first byte", () => {
    expect(run(fatal(), [bytes(0x41, 0x42, 0xE2), bytes(0x82)])).toEqual({ invalid: 2 })
    expect(run(fatal(), [bytes(0x41), bytes(0xF0), bytes(0x9F), bytes(0x98)])).toEqual({ invalid: 1 })
    expect(run(fatal(), [bytes(0x41, 0x42, 0x43, 0xF0, 0x9F), bytes(0x41)])).toEqual({ invalid: 3 })
    // an error found while the first bytes are still held back
    expect(run(fatal(), [bytes(0xC3)])).toEqual({ invalid: 0 })
  })

  it("fatal UTF-16 is checked by the platform, with no offset", () => {
    expect(run(fatal(), [bytes(0xFF, 0xFE, 0x41, 0x00, 0x42)])).toEqual({ invalid: undefined })
    expect(run(fatal(), [bytes(0xFE, 0xFF, 0x00), bytes(0x41)])).toEqual({ text: "A" })
    // a lone surrogate
    expect(run(fatal(), [bytes(0xFF, 0xFE, 0x00, 0xD8, 0x41, 0x00)])).toEqual({ invalid: undefined })
  })

  it("lenient decoding gives U+FFFD, never an error", () => {
    const input = new TextInput({ detect: "bom-only", encoding: "utf-8", fatal: false })
    expect(Result.merge(input.push(bytes(0x41, 0xFF, 0xC3)))).toBe("A\uFFFD")
    expect(Result.merge(input.end())).toBe("\uFFFD")
  })

  it("bom-only holds back 3 bytes, the longest byte-order mark", () => {
    const utf8 = fatal()
    expect(utf8.push(bytes(0x41, 0x42))).toEqual(Result.succeed(""))
    expect(utf8.push(bytes(0x43))).toEqual(Result.succeed("ABC"))
    const latin1 = fatal("windows-1252")
    expect(latin1.push(bytes(0x41, 0x42))).toEqual(Result.succeed(""))
    expect(latin1.push(bytes(0xE9))).toEqual(Result.succeed("ABé"))
    // the UTF-8 mark wins over the requested encoding, split or not
    const marked = fatal("windows-1252")
    expect(marked.push(bytes(0xEF, 0xBB))).toEqual(Result.succeed(""))
    expect(marked.push(bytes(0xBF, 0xC3, 0xA9))).toEqual(Result.succeed("é"))
  })

  it("detection holds back exactly DETECT_BYTES bytes", () => {
    const input = new TextInput({ detect: "bom-utf16-utf8-1252", fatal: false })
    expect(input.push(new Uint8Array(DETECT_BYTES - 1).fill(0x61))).toEqual(Result.succeed(""))
    expect(Result.merge(input.push(bytes(0x62))).length).toBe(DETECT_BYTES)
  })
})

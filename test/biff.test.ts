import fc from "fast-check"
import { describe, expect, it } from "vitest"
import { CONTINUE, RecordReader, readBof } from "../src/internal/biff/records.ts"
import { readRichString, SegmentReader } from "../src/internal/biff/segments.ts"
import { LimitExceeded, MalformedFile } from "../src/SpreadsheetError.ts"

const record = (type: number, data: ReadonlyArray<number> | Uint8Array): Uint8Array => {
  const out = new Uint8Array(4 + data.length)
  const v = new DataView(out.buffer)
  v.setUint16(0, type, true)
  v.setUint16(2, data.length, true)
  out.set(data, 4)
  return out
}
const concat = (parts: ReadonlyArray<Uint8Array>): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

describe("biff records", () => {
  it("reads records in order, each advancing by 4 + length; padding ends the walk", () => {
    const stream = concat([record(0x0809, [0, 6, 5, 0]), record(0x0000, []), record(0x00FC, [1, 2, 3]), record(0x000A, []), new Uint8Array([0, 0])])
    const reader = new RecordReader(stream, "Workbook")
    const seen: Array<[number, number, number]> = []
    let r = reader.next()
    while (r !== undefined && !(r instanceof MalformedFile)) {
      seen.push([r.type, r.offset, r.data.length])
      r = reader.next()
    }
    expect(r).toBeUndefined()
    expect(seen).toEqual([[0x0809, 0, 4], [0x0000, 8, 0], [0x00FC, 12, 3], [0x000A, 19, 0]])
    expect(reader.records).toBe(4)
    expect(readBof({ type: 0x0809, offset: 0, data: new Uint8Array([0, 6, 0x10, 0]) })).toEqual({ version: 0x0600, kind: 0x0010 })
    expect(readBof({ type: 0x0809, offset: 0, data: new Uint8Array([0, 6]) })).toBeUndefined()
  })

  it("a record running past the end is MalformedFile with its offset", () => {
    const stream = concat([record(0x0809, [0, 6, 5, 0]), new Uint8Array([0xFC, 0x00, 0xFF, 0x00, 1, 2])])
    const reader = new RecordReader(stream, "Workbook")
    reader.next()
    expect(reader.next()).toMatchObject({ _tag: "MalformedFile", part: "Workbook", offset: 8 })
  })

  it("any bytes: the walk ends after at most n/4 records, with records or a typed error", () => {
    fc.assert(
      fc.property(fc.uint8Array({ maxLength: 2000 }), (bytes) => {
        const reader = new RecordReader(bytes, "x")
        let steps = 0
        let r = reader.next()
        while (r !== undefined && !(r instanceof MalformedFile)) {
          steps++
          r = reader.next()
        }
        expect(steps).toBeLessThanOrEqual(bytes.length / 4)
      }),
      { numRuns: 1000 }
    )
  })

  it("gathers CONTINUE records with their record, within a byte limit", () => {
    const stream = concat([record(0x00FC, [1]), record(CONTINUE, [2, 3]), record(CONTINUE, []), record(CONTINUE, [4]), record(0x00FF, [9])])
    const reader = new RecordReader(stream, "Workbook")
    const sst = reader.nextWithContinues(100, "maxDecompressedBytes")
    expect(sst && !(sst instanceof MalformedFile) && !(sst instanceof LimitExceeded) && sst.segments.map((s) => Array.from(s))).toEqual([[1], [2, 3], [], [4]])
    expect(reader.peekType()).toBe(0x00FF)
    // 4 records' headers + 4 bytes = 20: 19 is over, under the caller's name
    const limited = new RecordReader(stream, "Workbook").nextWithContinues(19, "maxDecompressedBytes")
    expect(limited).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 19, actual: 20 })
    expect(new RecordReader(stream, "Workbook").nextWithContinues(20, "x")).toMatchObject({ type: 0x00FC })
  })

  it("empty CONTINUE records count against the limit (4 bytes each)", () => {
    const many = concat([record(0x00FC, [1]), ...Array.from({ length: 10_000 }, () => record(CONTINUE, []))])
    const reader = new RecordReader(many, "Workbook")
    const result = reader.nextWithContinues(1000, "maxDecompressedBytes")
    expect(result).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 1000 })
    // it stopped early: about 1000 / 4 records read, not 10 000
    expect(reader.records).toBeLessThanOrEqual(252)
  })
})

/** Encodes `text` as an SST entry split into segments at `cuts` (character positions), as Excel would. */
const splitString = (text: string, wide: boolean, cuts: ReadonlyArray<number>, extra: { runs?: number; ext?: number } = {}) => {
  const header = [text.length & 0xFF, text.length >> 8, (wide ? 1 : 0) | (extra.runs ? 0x08 : 0) | (extra.ext ? 0x04 : 0)]
  if (extra.runs) header.push(extra.runs & 0xFF, extra.runs >> 8)
  if (extra.ext) header.push(extra.ext & 0xFF, (extra.ext >> 8) & 0xFF, 0, 0)
  const points = [...new Set(cuts.filter((c) => c > 0 && c < text.length))].sort((a, b) => a - b)
  const segments: Array<Array<number>> = [header]
  let from = 0
  for (const point of [...points, text.length]) {
    if (from > 0) {
      // a continued string: a new segment starting with the option byte;
      // alternate the width to exercise the switch
      wide = !wide
      segments.push([wide ? 1 : 0])
    }
    const seg = segments[segments.length - 1]!
    for (let k = from; k < point; k++) {
      const c = text.charCodeAt(k)
      if (wide) seg.push(c & 0xFF, c >> 8)
      else seg.push(c & 0xFF)
    }
    from = point
  }
  // run and phonetic data after the characters, split once more if there is any
  const tail: Array<number> = []
  for (let k = 0; k < (extra.runs ?? 0) * 4 + (extra.ext ?? 0); k++) tail.push(0xAA)
  if (tail.length > 0) {
    segments[segments.length - 1]!.push(...tail.slice(0, 3))
    segments.push(tail.slice(3))
  }
  return segments.map((s) => new Uint8Array(s))
}

describe("biff segments", () => {
  it("numbers cross segment boundaries; reads past the end fail and stay failed", () => {
    const r = new SegmentReader([new Uint8Array([1]), new Uint8Array([2, 3]), new Uint8Array([]), new Uint8Array([4, 5])])
    expect(r.u16()).toBe(0x0201)
    expect(r.u16()).toBe(0x0403)
    expect(r.remaining()).toBe(1)
    expect(r.skip(1)).toBe(true)
    expect(r.remaining()).toBe(0)
    expect(r.u8()).toBeUndefined()
    expect(r.failed).toBe(true)
    expect(r.u8()).toBeUndefined()
    const s = new SegmentReader([new Uint8Array(3)])
    expect(s.skip(4)).toBe(false)
  })

  it("strings switch width at each CONTINUE (option byte), including right after the header", () => {
    // header ends its segment: the characters start with an option byte
    const segments = [new Uint8Array([3, 0, 0]), new Uint8Array([1, 0x41, 0, 0x42, 0, 0x43, 0])]
    expect(readRichString(new SegmentReader(segments), 100)).toBe("ABC")
    // 8-bit, then 16-bit, then 8-bit again
    expect(readRichString(new SegmentReader(splitString("héllo wörld", false, [3, 7])), 100)).toBe("héllo wörld")
    // 16-bit characters on both sides of a boundary
    const wide = (s: string) => s.split("").flatMap((c) => [c.charCodeAt(0) & 0xFF, c.charCodeAt(0) >> 8])
    const cjk = [new Uint8Array([4, 0, 1, ...wide("日本")]), new Uint8Array([1, ...wide("語!")])]
    expect(readRichString(new SegmentReader(cjk), 100)).toBe("日本語!")
  })

  it("8-bit characters are code units: 0x80-0x9F are C1 controls, not windows-1252", () => {
    const r = new SegmentReader([new Uint8Array([3, 0, 0, 0x80, 0x93, 0x9F])])
    expect(readRichString(r, 100)).toBe("\u0080\u0093\u009F")
  })

  it("remaining() is O(1): it doesn't walk the segments", () => {
    const parts = Array.from({ length: 1000 }, () => new Uint8Array([7]))
    let reads = 0
    const counted = new Proxy(parts, {
      get: (target, key, receiver) => {
        if (typeof key === "string" && /^\d+$/.test(key)) reads++
        return Reflect.get(target, key, receiver)
      }
    })
    const r = new SegmentReader(counted)
    const before = reads
    for (let k = 0; k < 1000; k++) r.remaining()
    expect(reads - before).toBe(0)
  })

  it("remaining() follows every kind of read", () => {
    const segments = splitString("abcdef", false, [2, 4], { runs: 1, ext: 2 })
    const total = segments.reduce((n, s) => n + s.length, 0)
    const r = new SegmentReader(segments)
    expect(r.remaining()).toBe(total)
    expect(readRichString(r, 100)).toBe("abcdef")
    expect(r.remaining()).toBe(0)
  })

  it("rich-text runs and phonetic data are skipped, across boundaries", () => {
    const segments = splitString("rich", false, [2], { runs: 2, ext: 5 })
    const reader = new SegmentReader([...segments, new Uint8Array([3, 0, 0, 0x78, 0x79, 0x7A])])
    expect(readRichString(reader, 100)).toBe("rich")
    expect(readRichString(reader, 100)).toBe("xyz")
    expect(reader.remaining()).toBe(0)
  })

  it("random strings split anywhere decode back", () => {
    fc.assert(
      fc.property(
        fc.string({ minLength: 1, maxLength: 60, unit: fc.constantFrom("a", "Z", "é", "ü", "ÿ", " ", "\u0080", "\u0085", "\u009F", "\u0000") }),
        fc.boolean(),
        fc.array(fc.nat(60), { maxLength: 6 }),
        (text, wide, cuts) => {
          expect(readRichString(new SegmentReader(splitString(text, wide, cuts)), 1000)).toBe(text)
        }
      ),
      { numRuns: 1000 }
    )
  })

  it("strings past maxLength, truncated strings and half characters fail", () => {
    expect(readRichString(new SegmentReader([new Uint8Array([200, 0, 0, ...new Uint8Array(200)])]), 100)).toBeNull()
    expect(readRichString(new SegmentReader([new Uint8Array([5, 0, 0, 0x41])]), 100)).toBeUndefined()
    expect(readRichString(new SegmentReader([new Uint8Array([2, 0, 1, 0x41])]), 100)).toBeUndefined()
    // a huge phonetic block size can't skip past the end
    const huge = new SegmentReader([new Uint8Array([1, 0, 4, 0xFF, 0xFF, 0xFF, 0x7F, 0x41])])
    expect(readRichString(huge, 100)).toBeUndefined()
  })
})

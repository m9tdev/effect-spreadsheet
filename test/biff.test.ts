import fc from "fast-check"
import { describe, expect, it } from "vitest"
import { readRichString, RecordCursor } from "../src/internal/biff/cursor.ts"
import { type BiffRecord, CONTINUE, Embedded, RecordReader, readBof, SubstreamReader } from "../src/internal/biff/records.ts"
import type { StringRead } from "../src/internal/binary/layout.ts"
import { MalformedFile } from "../src/SpreadsheetError.ts"

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

/** The record the reader reads next (it must be one). */
const nextOf = (reader: RecordReader): BiffRecord => {
  const r = reader.next()
  if (r === undefined || r instanceof MalformedFile) throw new Error("expected a record")
  return r
}

/** Every byte a cursor reads, one u8 at a time, up to the end of its record's data. */
const bytesOf = (cursor: RecordCursor): Array<number> => {
  const out: Array<number> = []
  // progress: each u8 consumes a byte that has(1) found
  while (cursor.has(1)) out.push(cursor.u8())
  return out
}

/**
 * A cursor over `parts`: the first is a record's data, each other one a
 * CONTINUE's (an empty one an empty CONTINUE), then an EOF that isn't part
 * of them.
 */
const cursorOf = (parts: ReadonlyArray<Uint8Array>): RecordCursor => {
  const stream = concat([record(0x00FC, parts[0] ?? []), ...parts.slice(1).map((p) => record(CONTINUE, p)), record(0x000A, [])])
  return new RecordCursor(stream, { offset: 0, data: stream.subarray(4, 4 + (parts[0]?.length ?? 0)) }, stream.length)
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

  it("passes over the CONTINUE records after a record, within a byte limit; a cursor reads their data with the record's", () => {
    const stream = concat([record(0x00FC, [1]), record(CONTINUE, [2, 3]), record(CONTINUE, []), record(CONTINUE, [4]), record(0x00FF, [9])])
    const reader = new RecordReader(stream, "Workbook")
    const sst = nextOf(reader)
    expect(reader.continues(sst, Infinity, 100, "maxDecompressedBytes")).toBe("done")
    expect(reader.peekType()).toBe(0x00FF)
    // the empty CONTINUE continues nothing
    expect(bytesOf(new RecordCursor(stream, sst, reader.offset))).toEqual([1, 2, 3, 4])
    // 4 records' headers + 4 bytes = 20: 19 is over, under the caller's name
    const limited = new RecordReader(stream, "Workbook")
    expect(limited.continues(nextOf(limited), Infinity, 19, "maxDecompressedBytes"))
      .toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 19, actual: 20 })
    const fits = new RecordReader(stream, "Workbook")
    expect(fits.continues(nextOf(fits), Infinity, 20, "x")).toBe("done")
  })

  it("10 million empty CONTINUEs (40 MB) under a 1 GiB cap: passed over, and a cursor reads the record's one byte", () => {
    const stream = new Uint8Array(4 + 1 + 10_000_000 * 4)
    const v = new DataView(stream.buffer)
    v.setUint16(0, 0x00FC, true)
    v.setUint16(2, 1, true)
    stream[4] = 7
    for (let k = 0; k < 10_000_000; k++) v.setUint16(5 + 4 * k, CONTINUE, true)
    const reader = new RecordReader(stream, "Workbook")
    const first = nextOf(reader)
    expect(reader.continues(first, Infinity, 1024 * 1024 * 1024, "maxDecompressedBytes")).toBe("done")
    expect(reader.records).toBe(10_000_001)
    expect(reader.offset).toBe(stream.length)
    expect(bytesOf(new RecordCursor(stream, first, reader.offset))).toEqual([7])
  })

  it("1-byte CONTINUEs: a cursor reads across each of them, in order", () => {
    const n = 1_000_000
    const stream = new Uint8Array(4 + 1 + n * 5)
    const v = new DataView(stream.buffer)
    v.setUint16(0, 0x00FC, true)
    v.setUint16(2, 1, true)
    for (let k = 0; k < n; k++) {
      v.setUint16(5 + 5 * k, CONTINUE, true)
      v.setUint16(7 + 5 * k, 1, true)
      stream[9 + 5 * k] = k & 0xFF
    }
    const reader = new RecordReader(stream, "Workbook")
    const first = nextOf(reader)
    expect(reader.continues(first, Infinity, 1024 * 1024 * 1024, "maxDecompressedBytes")).toBe("done")
    const r = new RecordCursor(stream, first, reader.offset)
    expect(r.has(n + 1)).toBe(true)
    expect(r.has(n + 2)).toBe(false)
    r.skip(1)
    expect([r.u8(), r.u8(), r.u8()]).toEqual([0, 1, 2])
    expect(r.u32()).toBe(0x06050403)
    expect(r.skip(n - 9)).toBe(true)
    expect(r.consumed()).toBe(n - 1)
    expect([r.u8(), r.u8(), r.has(1), r.u8()]).toEqual([(n - 2) & 0xFF, (n - 1) & 0xFF, false, 0])
  })

  it("empty CONTINUE records count against the limit (4 bytes each)", () => {
    const many = concat([record(0x00FC, [1]), ...Array.from({ length: 10_000 }, () => record(CONTINUE, []))])
    const reader = new RecordReader(many, "Workbook")
    const result = reader.continues(nextOf(reader), Infinity, 1000, "maxDecompressedBytes")
    expect(result).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 1000, actual: 1001 })
    // it stopped early: about 1000 / 4 records read, not 10 000
    expect(reader.records).toBeLessThanOrEqual(252)
  })
})

describe("biff substream reader", () => {
  /** Each call's result: [type or "embedded"/"opens", offset], up to the end or an error, and the records counted. */
  const walk = (stream: Uint8Array) => {
    const reader = new SubstreamReader(stream, "sheet")
    const seen: Array<[number | string, number]> = []
    // progress: every read() reads one record (at most stream.length / 4 of them)
    for (let k = 0; k <= stream.length / 4; k++) {
      const r = reader.read()
      if (r === undefined) return { seen, end: "end", records: reader.records }
      if (r instanceof MalformedFile) return { seen, end: r, records: reader.records }
      seen.push(r instanceof Embedded ? [r.opens ? "opens" : "embedded", r.offset] : [r.type, r.offset])
    }
    throw new Error("the reader didn't advance")
  }
  const BOF = 0x0809
  const EOF = 0x000A
  const NUMBER = 0x0203

  it("an embedded substream, nested ones too, comes as one opening marker, then skipped records up to its EOF", () => {
    const stream = concat([
      record(NUMBER, [1]), // 0
      record(BOF, [0, 6, 0x20, 0]), // 5
      record(NUMBER, [2]), // 13
      record(BOF, [0, 6, 0x20, 0]), // 18
      record(EOF, []), // 26
      record(NUMBER, [3]), // 30
      record(EOF, []), // 35
      record(NUMBER, [4]), // 39
      record(BOF, [0, 6, 0x20, 0]), // 44
      record(EOF, []), // 52
      record(EOF, []) // 56: the substream's own EOF
    ])
    expect(walk(stream)).toEqual({
      seen: [
        [NUMBER, 0],
        ["opens", 5],
        ["embedded", 13],
        ["embedded", 18],
        ["embedded", 26],
        ["embedded", 30],
        ["embedded", 35],
        [NUMBER, 39],
        ["opens", 44],
        ["embedded", 52],
        [EOF, 56]
      ],
      end: "end",
      records: 11
    })
  })

  it("a BIFF2-4 BOF is a record like any other, inside an embedded substream too", () => {
    const stream = concat([
      record(0x0409, [0, 0, 0x20, 0]), // 0
      record(BOF, [0, 6, 0x20, 0]), // 8
      record(0x0209, [0, 0, 0x20, 0]), // 16
      record(EOF, []), // 24
      record(0x0009, [0, 0, 0x20, 0]), // 28
      record(EOF, []) // 36
    ])
    expect(walk(stream).seen).toEqual([[0x0409, 0], ["opens", 8], ["embedded", 16], ["embedded", 24], [0x0009, 28], [EOF, 36]])
  })

  it("an unterminated embedded substream runs to the end; a record past it is the reader's error", () => {
    expect(walk(concat([record(BOF, [0, 6, 0x20, 0]), record(NUMBER, [1])])).seen).toEqual([["opens", 0], ["embedded", 8]])
    const cut = walk(concat([record(BOF, [0, 6, 0x20, 0]), new Uint8Array([0x03, 0x02, 0x10, 0x00, 1])]))
    expect(cut.seen).toEqual([["opens", 0]])
    expect(cut.end).toMatchObject({ _tag: "MalformedFile", offset: 8 })
  })

  it("a record of the continued type comes with its CONTINUEs, only in the substream itself", () => {
    const STRING = 0x0207
    const stream = concat([
      record(STRING, [1]), // 0
      record(CONTINUE, [2]), // 5
      record(CONTINUE, []), // 10
      record(BOF, [0, 6, 0x20, 0]), // 14
      record(STRING, [3]), // 22
      record(CONTINUE, [4]), // 27
      record(EOF, []), // 32
      record(NUMBER, [5]), // 36
      record(CONTINUE, [6]) // 41
    ])
    const reader = new SubstreamReader(stream, "sheet", 0, stream.length, STRING)
    const first = reader.read()
    if (first === undefined || first instanceof MalformedFile || first instanceof Embedded) throw new Error("expected a record")
    expect([first.type, first.offset, Array.from(first.data)]).toEqual([STRING, 0, [1]])
    expect([reader.offset, reader.records]).toEqual([14, 3])
    expect(bytesOf(new RecordCursor(stream, first, stream.length))).toEqual([1, 2])
    expect(reader.read()).toEqual(new Embedded(14, true))
    // inside the embedded substream, a STRING is skipped as any record, and so is its CONTINUE
    expect(reader.read()).toEqual(new Embedded(22, false))
    expect(reader.read()).toEqual(new Embedded(27, false))
    expect(reader.read()).toEqual(new Embedded(32, false))
    // another type's CONTINUE is a record of its own
    expect(reader.read()).toMatchObject({ type: NUMBER, offset: 36 })
    expect(reader.read()).toMatchObject({ type: CONTINUE, offset: 41 })
    expect(reader.read()).toBeUndefined()
    // without a continued type, the STRING's CONTINUEs are records too
    expect(walk(stream).seen.slice(0, 3)).toEqual([[STRING, 0], [CONTINUE, 5], [CONTINUE, 10]])
  })

  it("a CONTINUE of the continued type's record that runs past the end is MalformedFile with its offset", () => {
    const STRING = 0x0207
    const stream = concat([record(STRING, [1]), record(CONTINUE, [2]), new Uint8Array([0x3C, 0x00, 0x09, 0x00, 1])])
    const reader = new SubstreamReader(stream, "sheet", 0, stream.length, STRING)
    expect(reader.read()).toMatchObject({ _tag: "MalformedFile", offset: 10 })
  })
})

describe("biff CONTINUE records in bounded steps", () => {
  // an SST with three CONTINUEs (one of them empty), then an EOF
  const stream = concat([record(0x00FC, [1]), record(CONTINUE, [2, 3]), record(CONTINUE, []), record(CONTINUE, [4]), record(0x000A, [])])
  const start = (bytes: Uint8Array) => {
    const reader = new RecordReader(bytes, "Workbook")
    return { reader, first: nextOf(reader) }
  }

  it("a step that reaches its bound before the last CONTINUE is more; a later one is done, with every segment read by a cursor", () => {
    const { first, reader } = start(stream)
    expect(reader.continues(first, 1, 100, "maxDecompressedBytes")).toBe("more")
    expect(reader.continues(first, 1, 100, "maxDecompressedBytes")).toBe("more")
    expect(reader.records).toBe(3)
    expect(reader.continues(first, 1, 100, "maxDecompressedBytes")).toBe("done")
    expect(reader.peekType()).toBe(0x000A)
    expect(bytesOf(new RecordCursor(stream, first, reader.offset))).toEqual([1, 2, 3, 4])
  })

  it("a step after done is done again and reads nothing", () => {
    const { first, reader } = start(stream)
    expect(reader.continues(first, 10, 100, "maxDecompressedBytes")).toBe("done")
    const records = reader.records
    const offset = reader.offset
    expect(reader.continues(first, 10, 100, "maxDecompressedBytes")).toBe("done")
    expect([reader.records, reader.offset]).toEqual([records, offset])
  })

  it("a step that goes past the byte limit is LimitExceeded with the size so far, after some more steps", () => {
    // 4 records' headers + 4 bytes = 20: the last CONTINUE goes over 19
    const { first, reader } = start(stream)
    expect(reader.continues(first, 1, 19, "maxDecompressedBytes")).toBe("more")
    expect(reader.continues(first, 1, 19, "maxDecompressedBytes")).toBe("more")
    expect(reader.continues(first, 1, 19, "maxDecompressedBytes")).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 19, actual: 20 })
  })

  it("an empty CONTINUE counts towards the limit but adds no data", () => {
    const empties = concat([record(0x00FC, [1]), record(CONTINUE, []), record(CONTINUE, [])])
    // two empty CONTINUEs: 5 + 4 + 4 = 13 bytes
    const fits = start(empties)
    expect(fits.reader.continues(fits.first, 10, 13, "maxDecompressedBytes")).toBe("done")
    expect(bytesOf(new RecordCursor(empties, fits.first, fits.reader.offset))).toEqual([1])
    const over = start(empties)
    expect(over.reader.continues(over.first, 10, 12, "maxDecompressedBytes")).toMatchObject({ _tag: "LimitExceeded", max: 12, actual: 13 })
  })

  it("a CONTINUE that runs past the end of the stream is MalformedFile", () => {
    const cut = concat([record(0x00FC, [1]), record(CONTINUE, [2, 3])]).subarray(0, 10)
    const { first, reader } = start(cut)
    expect(reader.continues(first, 10, 100, "maxDecompressedBytes")).toMatchObject({ _tag: "MalformedFile", offset: 5 })
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

/** A string read's text, or the failure itself. */
const textOf = (read: StringRead) => (read._tag === "Ok" ? read.value : read)

describe("biff record cursor", () => {
  it("numbers cross segment boundaries; has() checks what is left; skips past the end fail", () => {
    const r = cursorOf([new Uint8Array([1]), new Uint8Array([2, 3]), new Uint8Array([]), new Uint8Array([4, 5])])
    expect(r.has(5)).toBe(true)
    expect(r.has(6)).toBe(false)
    expect(r.u16()).toBe(0x0201)
    expect(r.u16()).toBe(0x0403)
    expect(r.consumed()).toBe(4)
    expect(r.has(1)).toBe(true)
    expect(r.has(2)).toBe(false)
    expect(r.skip(1)).toBe(true)
    expect(r.consumed()).toBe(5)
    expect(r.has(1)).toBe(false)
    const s = cursorOf([new Uint8Array(3)])
    expect(s.skip(4)).toBe(false)
    expect([s.consumed(), s.has(3)]).toEqual([0, true])
    const u = cursorOf([new Uint8Array([0x78, 0x56]), new Uint8Array([0x34, 0x12])])
    expect(u.u32()).toBe(0x12345678)
  })

  it("strings switch width at each CONTINUE (option byte), including right after the header", () => {
    // header ends its segment: the characters start with an option byte
    const segments = [new Uint8Array([3, 0, 0]), new Uint8Array([1, 0x41, 0, 0x42, 0, 0x43, 0])]
    expect(textOf(readRichString(cursorOf(segments), 100))).toBe("ABC")
    // 8-bit, then 16-bit, then 8-bit again
    expect(textOf(readRichString(cursorOf(splitString("héllo wörld", false, [3, 7])), 100))).toBe("héllo wörld")
    // 16-bit characters on both sides of a boundary
    const wide = (s: string) => s.split("").flatMap((c) => [c.charCodeAt(0) & 0xFF, c.charCodeAt(0) >> 8])
    const cjk = [new Uint8Array([4, 0, 1, ...wide("日本")]), new Uint8Array([1, ...wide("語!")])]
    expect(textOf(readRichString(cursorOf(cjk), 100))).toBe("日本語!")
  })

  it("a segment holding only an option byte is followed by the next segment's option byte, not a character", () => {
    // 8-bit: "ab" | [00] | [00 "c"], then another string
    const narrow = cursorOf([
      new Uint8Array([3, 0, 0, 0x61, 0x62]),
      new Uint8Array([0]),
      new Uint8Array([0, 0x63, 3, 0, 0, 0x78, 0x79, 0x7A])
    ])
    expect(textOf(readRichString(narrow, 100))).toBe("abc")
    expect(textOf(readRichString(narrow, 100))).toBe("xyz")
    expect(narrow.has(1)).toBe(false)
    // 16-bit: "a" | [01] | [01 "b" "c"]
    const wide = cursorOf([new Uint8Array([3, 0, 0, 0x61]), new Uint8Array([1]), new Uint8Array([1, 0x62, 0, 0x63, 0])])
    expect(textOf(readRichString(wide, 100))).toBe("abc")
    // several in a row, right after the header, widths switching: the last option byte wins
    const many = cursorOf([
      new Uint8Array([2, 0, 0]),
      new Uint8Array([1]),
      new Uint8Array([0]),
      new Uint8Array([1]),
      new Uint8Array([1, 0xE5, 0x65, 0x2C, 0x67])
    ])
    expect(textOf(readRichString(many, 100))).toBe("日本")
    // option bytes and then nothing left: truncated
    expect(textOf(readRichString(cursorOf([new Uint8Array([2, 0, 0, 0x61]), new Uint8Array([0]), new Uint8Array([0])]), 100)))
      .toEqual({ _tag: "Truncated" })
  })

  it("8-bit characters are code units: 0x80-0x9F are C1 controls, not windows-1252", () => {
    const r = cursorOf([new Uint8Array([3, 0, 0, 0x80, 0x93, 0x9F])])
    expect(textOf(readRichString(r, 100))).toBe("\u0080\u0093\u009F")
  })

  it("has() looks ahead across CONTINUEs without moving; reads then go on from where they were", () => {
    const r = cursorOf([new Uint8Array([1, 2]), new Uint8Array([]), new Uint8Array([3]), new Uint8Array([4, 5, 6])])
    expect(r.u8()).toBe(1)
    expect(r.has(5)).toBe(true)
    expect(r.has(6)).toBe(false)
    expect(r.consumed()).toBe(1)
    expect([r.u16(), r.u8()]).toEqual([0x0302, 4])
    expect([r.has(2), r.has(3)]).toEqual([true, false])
    expect(r.skip(3)).toBe(false)
    expect(r.u16()).toBe(0x0605)
    expect([r.has(1), r.u8(), r.consumed()]).toEqual([false, 0, 6])
  })

  it("the record's data ends at the first record that isn't a CONTINUE, or one that doesn't fit before the end", () => {
    const stream = concat([record(0x00FC, [1]), record(CONTINUE, [2]), record(0x00FF, [3]), record(CONTINUE, [4])])
    const first = { offset: 0, data: stream.subarray(4, 5) }
    expect(bytesOf(new RecordCursor(stream, first, stream.length))).toEqual([1, 2])
    // a CONTINUE whose header fits before the end, but not its data
    expect(bytesOf(new RecordCursor(stream, first, 9))).toEqual([1])
    expect(bytesOf(new RecordCursor(stream, first, 10))).toEqual([1, 2])
    // an end past the stream reads only the stream
    const cut = stream.subarray(0, 9)
    expect(bytesOf(new RecordCursor(cut, first, 100))).toEqual([1])
    // with the end at the record's own end, a cursor over one range
    expect(bytesOf(new RecordCursor(stream, first, 5))).toEqual([1])
  })

  it("characters at the very start of the record's data have no option byte; those at a CONTINUE's start do", () => {
    expect(textOf(cursorOf([new Uint8Array([0x41, 0x42])]).chars(2, false))).toBe("AB")
    expect(textOf(cursorOf([new Uint8Array([]), new Uint8Array([0, 0x41, 0x42])]).chars(2, false))).toBe("AB")
  })

  it("consumed() follows every kind of read", () => {
    const segments = splitString("abcdef", false, [2, 4], { runs: 1, ext: 2 })
    const total = segments.reduce((n, s) => n + s.length, 0)
    const r = cursorOf(segments)
    expect(r.has(total)).toBe(true)
    expect(r.has(total + 1)).toBe(false)
    const read = readRichString(r, 100)
    expect(textOf(read)).toBe("abcdef")
    expect([r.consumed(), read._tag === "Ok" && read.end]).toEqual([total, total])
    expect(r.has(1)).toBe(false)
  })

  it("rich-text runs and phonetic data are skipped, across boundaries", () => {
    const segments = splitString("rich", false, [2], { runs: 2, ext: 5 })
    const reader = cursorOf([...segments, new Uint8Array([3, 0, 0, 0x78, 0x79, 0x7A])])
    expect(textOf(readRichString(reader, 100))).toBe("rich")
    expect(textOf(readRichString(reader, 100))).toBe("xyz")
    expect(reader.has(1)).toBe(false)
  })

  it("random strings split anywhere decode back", () => {
    fc.assert(
      fc.property(
        fc.string({ minLength: 1, maxLength: 60, unit: fc.constantFrom("a", "Z", "é", "ü", "ÿ", " ", "\u0080", "\u0085", "\u009F", "\u0000") }),
        fc.boolean(),
        fc.array(fc.nat(60), { maxLength: 6 }),
        (text, wide, cuts) => {
          expect(textOf(readRichString(cursorOf(splitString(text, wide, cuts)), 1000))).toBe(text)
        }
      ),
      { numRuns: 1000 }
    )
  })

  it("strings past maxLength, truncated strings and half characters fail", () => {
    expect(textOf(readRichString(cursorOf([new Uint8Array([200, 0, 0, ...new Uint8Array(200)])]), 100))).toEqual({ _tag: "TooLong", count: 200 })
    expect(textOf(readRichString(cursorOf([new Uint8Array([5, 0, 0, 0x41])]), 100))).toEqual({ _tag: "Truncated" })
    expect(textOf(readRichString(cursorOf([new Uint8Array([2, 0, 1, 0x41])]), 100))).toEqual({ _tag: "Truncated" })
    // a 16-bit character split across a CONTINUE boundary: its first byte ends a segment
    expect(textOf(readRichString(cursorOf([new Uint8Array([2, 0, 1, 0x41, 0x00, 0x42]), new Uint8Array([1, 0x00, 0x43, 0x00])]), 100)))
      .toEqual({ _tag: "Truncated" })
    // a huge phonetic block size can't skip past the end
    const huge = cursorOf([new Uint8Array([1, 0, 4, 0xFF, 0xFF, 0xFF, 0x7F, 0x41])])
    expect(textOf(readRichString(huge, 100))).toEqual({ _tag: "Truncated" })
  })
})

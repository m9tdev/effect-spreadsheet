import fc from "fast-check"
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { readNullableWideString, readRichStr, readWideString } from "../src/internal/biff12/fields.ts"
import { type Biff12Record, readRecords, RecordSplitter, splitAll } from "../src/internal/biff12/records.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { concat, nullableWide, record, u32, varint, wide } from "./helpers/xlsb.ts"

const MAX = SpreadsheetLimits.defaults.maxRecordBytes

const summary = (records: ReadonlyArray<Biff12Record>) => records.map((r) => [r.type, r.offset, Array.from(r.data)])

/** Feeds `bytes` cut at `cuts`; the records, or the first error. */
const feedChunks = (bytes: Uint8Array, cuts: ReadonlyArray<number>, max = MAX): Array<Biff12Record> | SpreadsheetError => {
  const splitter = new RecordSplitter("test.bin", max)
  const out: Array<Biff12Record> = []
  const points = [...new Set(cuts.filter((c) => c > 0 && c < bytes.length))].sort((a, b) => a - b)
  let from = 0
  for (const cut of [...points, bytes.length]) {
    const error = splitter.feed(bytes.subarray(from, cut), (r) => {
      // copy: a view of a chunk is only valid while the chunk is
      out.push({ ...r, data: r.data.slice() })
      return undefined
    })
    if (error !== undefined) return error
    from = cut
  }
  return splitter.end() ?? out
}

describe("biff12 records", () => {
  it("encodes types in 1-2 bytes and sizes in 1-4 bytes, low bits first", () => {
    expect(varint(0)).toEqual([0])
    expect(varint(127)).toEqual([127])
    expect(varint(128)).toEqual([0x80, 1])
    expect(varint(0x3FFF)).toEqual([0xFF, 0x7F])
    expect(varint(2 ** 28 - 1)).toEqual([0xFF, 0xFF, 0xFF, 0x7F])
  })

  it("splits records in order, with their header offsets; empty records too", () => {
    const bytes = concat([record(0x83), record(0x9C, [1, 2, 3]), record(0, []), record(0x0417, new Array(200).fill(7)), record(5, [9])])
    const records = splitAll(bytes, "test.bin", MAX)
    expect(Array.isArray(records)).toBe(true)
    expect((records as Array<Biff12Record>).map((r) => [r.type, r.offset, r.data.length])).toEqual([
      [0x83, 0, 0],
      [0x9C, 3, 3],
      [0, 9, 0],
      [0x0417, 11, 200],
      [5, 215, 1]
    ])
  })

  it("gives the same records however the part is cut into chunks", () => {
    const arbRecord = fc.record({
      type: fc.integer({ min: 0, max: 0x3FFF }),
      data: fc.uint8Array({ maxLength: 300 })
    })
    fc.assert(
      fc.property(fc.array(arbRecord, { maxLength: 20 }), fc.array(fc.nat(7000), { maxLength: 40 }), (records, cuts) => {
        const bytes = concat(records.map((r) => record(r.type, r.data)))
        const whole = splitAll(bytes, "test.bin", MAX)
        const chunked = feedChunks(bytes, cuts)
        expect(Array.isArray(whole) && Array.isArray(chunked)).toBe(true)
        expect(summary(chunked as Array<Biff12Record>)).toEqual(summary(whole as Array<Biff12Record>))
        expect((whole as Array<Biff12Record>).map((r) => [r.type, Array.from(r.data)])).toEqual(records.map((r) => [r.type, Array.from(r.data)]))
      }),
      { numRuns: 300 }
    )
  })

  it("a record inside one chunk is a view of it, not a copy", () => {
    const chunk = record(7, [1, 2, 3, 4])
    let seen: Uint8Array | undefined
    new RecordSplitter("test.bin", MAX).feed(chunk, (r) => {
      seen = r.data
      return undefined
    })
    expect(seen!.buffer).toBe(chunk.buffer)
  })

  it("rejects a type longer than 2 bytes and a size longer than 4 bytes, with the record's offset", () => {
    const longType = concat([record(1, [0]), new Uint8Array([0x80, 0x80, 0x01, 0x00])])
    expect(splitAll(longType, "t.bin", MAX)).toMatchObject({ _tag: "MalformedFile", format: "xlsb", part: "t.bin", offset: 3, reason: "record type longer than 2 bytes" })
    const longSize = new Uint8Array([0x05, 0x80, 0x80, 0x80, 0x80, 0x01])
    expect(splitAll(longSize, "t.bin", MAX)).toMatchObject({ _tag: "MalformedFile", offset: 0, reason: "record 0x5: size longer than 4 bytes" })
  })

  it("checks the size against maxRecordBytes from the header alone, before any data", () => {
    // a header declaring 2^28 - 1 bytes, with no data after it
    const header = new Uint8Array([0x13, 0xFF, 0xFF, 0xFF, 0x7F])
    const splitter = new RecordSplitter("t.bin", MAX)
    const error = splitter.feed(header, () => undefined)
    expect(error).toBeInstanceOf(LimitExceeded)
    expect(error).toMatchObject({ limit: "maxRecordBytes", max: MAX, actual: 2 ** 28 - 1 })
    // at the limit is fine
    expect(splitAll(record(1, new Uint8Array(100)), "t.bin", 100)).toHaveLength(1)
    expect(splitAll(record(1, new Uint8Array(101)), "t.bin", 100)).toMatchObject({ _tag: "LimitExceeded", actual: 101 })
  })

  it("a record cut off by the end of the part is MalformedFile, header or body", () => {
    const bytes = concat([record(1, [1]), record(0x9C, [1, 2, 3, 4, 5])])
    expect(splitAll(bytes.subarray(0, 4), "t.bin", MAX)).toMatchObject({ _tag: "MalformedFile", offset: 3, reason: "record header cut off by the end of the part" })
    expect(splitAll(bytes.subarray(0, 8), "t.bin", MAX)).toMatchObject({
      _tag: "MalformedFile",
      offset: 3,
      reason: "record 0x9c runs past the end of the part (2 of 5 bytes)"
    })
    // a 2-byte type cut after its first byte
    expect(splitAll(new Uint8Array([0x83]), "t.bin", MAX)).toBeInstanceOf(MalformedFile)
    expect(splitAll(new Uint8Array(0), "t.bin", MAX)).toEqual([])
  })

  it("stops at the handler's error and returns it", () => {
    const bytes = concat([record(1), record(2), record(3)])
    const seen: Array<number> = []
    const stop = new MalformedFile({ format: "xlsb", reason: "stop" })
    const error = new RecordSplitter("t.bin", MAX).feed(bytes, (r) => {
      seen.push(r.type)
      return r.type === 2 ? stop : undefined
    })
    expect(error).toBe(stop)
    expect(seen).toEqual([1, 2])
  })

  it("a big record fed one byte at a time is copied once: work stays linear", () => {
    const size = 1 << 20
    const bytes = record(19, new Uint8Array(size).fill(0xAB))
    const splitter = new RecordSplitter("t.bin", MAX)
    let got: Uint8Array | undefined
    const t0 = performance.now()
    // progress: at up to bytes.length
    for (let at = 0; at < bytes.length; at++) {
      splitter.feed(bytes.subarray(at, at + 1), (r) => {
        got = r.data
        return undefined
      })
    }
    const elapsed = performance.now() - t0
    expect(splitter.end()).toBeUndefined()
    expect(got?.length).toBe(size)
    expect(got![size - 1]).toBe(0xAB)
    // a copy of the record per byte would be 2^40 byte copies; once is ~1M steps (generous bound for loaded machines)
    expect(elapsed).toBeLessThan(10_000)
  })

  it("n bytes give at most n / 2 records, and every byte is consumed", () => {
    fc.assert(
      fc.property(fc.uint8Array({ maxLength: 2000 }), (bytes) => {
        const splitter = new RecordSplitter("t.bin", 1000)
        let records = 0
        const error = splitter.feed(bytes, () => {
          records++
          return undefined
        })
        expect(records).toBeLessThanOrEqual(bytes.length / 2)
        if (error === undefined) expect(splitter.offset).toBe(bytes.length)
      }),
      { numRuns: 500 }
    )
  })

  it("readRecords reads a stream of chunks, and fails on a cut-off record or a stream error", async () => {
    const bytes = concat([record(0x83), record(0x9C, [1, 2, 3]), record(0x84)])
    const seen: Array<number> = []
    const stats = await Effect.runPromise(readRecords(Stream.make(bytes.subarray(0, 4), bytes.subarray(4)), (r) => {
      seen.push(r.type)
      return undefined
    }, { maxRecordBytes: MAX }, "t.bin"))
    expect(stats.records).toBe(3)
    expect(seen).toEqual([0x83, 0x9C, 0x84])
    const cut = await Effect.runPromiseExit(readRecords(Stream.make(bytes.subarray(0, 6)), () => undefined, { maxRecordBytes: MAX }, "t.bin"))
    expect(Exit.isFailure(cut) && JSON.stringify(cut.cause)).toContain("runs past the end of the part")
    const failing = Stream.concat(Stream.make(bytes), Stream.fail(new MalformedFile({ format: "zip", reason: "CRC-32 mismatch" })))
    const exit = await Effect.runPromiseExit(readRecords(failing, () => undefined, { maxRecordBytes: MAX }, "t.bin"))
    expect(Exit.isFailure(exit) && JSON.stringify(exit.cause)).toContain("CRC-32 mismatch")
  })
})

describe("biff12 strings", () => {
  it("reads an XLWideString and the offset after it", () => {
    const data = new Uint8Array([0xEE, ...wide("Sheet1"), 0xDD])
    expect(readWideString(data, 1, 100)).toEqual({ value: "Sheet1", end: 1 + 4 + 12 })
    expect(readWideString(new Uint8Array(wide("")), 0, 100)).toEqual({ value: "", end: 4 })
    expect(readWideString(new Uint8Array(wide("é😀")), 0, 100)).toEqual({ value: "é😀", end: 4 + 6 })
  })

  it("past the data is undefined; a count over the limit is null, before decoding", () => {
    const data = new Uint8Array(wide("abcdef"))
    expect(readWideString(data.subarray(0, 15), 0, 100)).toBeUndefined()
    expect(readWideString(data, 2, Infinity)).toBeUndefined()
    expect(readWideString(data, -1, 100)).toBeUndefined()
    expect(readWideString(data, 0, 5)).toBeNull()
    expect(readWideString(data, 0, 6)).toEqual({ value: "abcdef", end: 16 })
    // a huge count in a tiny record: null under a small limit, undefined (past the data) under Infinity
    const huge = new Uint8Array(u32(0x7FFFFFFF))
    expect(readWideString(huge, 0, 1000)).toBeNull()
    expect(readWideString(huge, 0, Infinity)).toBeUndefined()
  })

  it("an XLNullableWideString of count 0xFFFFFFFF is no string", () => {
    expect(readNullableWideString(new Uint8Array(nullableWide(undefined)), 0, 100)).toEqual({ value: undefined, end: 4 })
    expect(readNullableWideString(new Uint8Array(nullableWide("rId1")), 0, 100)).toEqual({ value: "rId1", end: 12 })
    expect(readNullableWideString(new Uint8Array(3), 0, 100)).toBeUndefined()
  })

  it("a RichStr's text follows its flags byte; runs and phonetic data after it are ignored", () => {
    const rich = new Uint8Array([0x03, ...wide("hello"), ...u32(1), 0, 0, 1, 0, 9, 9, 9])
    expect(readRichStr(rich, 0, 100)).toEqual({ value: "hello", end: 15 })
    expect(readRichStr(new Uint8Array(0), 0, 100)).toBeUndefined()
    expect(readRichStr(rich, 0, 4)).toBeNull()
  })
})

import fc from "fast-check"
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { readNullableWideString, readRichStr, readWideString } from "../src/internal/biff12/fields.ts"
import { type Biff12Record, readRecords, RecordSplitter } from "../src/internal/biff12/records.ts"
import { FixedReader } from "../src/internal/binary/layout.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { counter, metered } from "./helpers/meter.ts"
import { concat, f64, nullableWide, record, u32, varint, wide } from "./helpers/xlsb.ts"

const MAX = SpreadsheetLimits.defaults.maxRecordBytes

/** Every record of a whole part, or the first error. */
const splitAll = (bytes: Uint8Array, part: string, maxRecordBytes: number): Array<Biff12Record> | SpreadsheetError => {
  const out: Array<Biff12Record> = []
  const splitter = new RecordSplitter(part, maxRecordBytes)
  const error = splitter.feed(bytes, (record) => {
    out.push(record)
    return undefined
  }) ?? splitter.end()
  return error ?? out
}

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

  it("a header declaring 2^28 - 1 bytes allocates nothing: buffers grow by under 1 MiB", () => {
    const header = new Uint8Array([0x13, 0xFF, 0xFF, 0xFF, 0x7F])
    const before = process.memoryUsage().arrayBuffers
    const splitter = new RecordSplitter("t.bin", MAX)
    const error = splitter.feed(header, () => undefined)
    const grown = process.memoryUsage().arrayBuffers - before
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxRecordBytes" })
    // a body allocated before the check would be 256 MiB, still held by the splitter here
    expect(grown).toBeLessThan(1 << 20)
    expect(splitter.records).toBe(0)
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
    const records = counter("records")
    await Effect.runPromise(metered(records, readRecords(Stream.make(bytes.subarray(0, 4), bytes.subarray(4)), (r) => {
      seen.push(r.type)
      return undefined
    }, { maxRecordBytes: MAX }, "t.bin")))
    expect(records.total()).toBe(3)
    expect(seen).toEqual([0x83, 0x9C, 0x84])
    const cut = await Effect.runPromiseExit(readRecords(Stream.make(bytes.subarray(0, 6)), () => undefined, { maxRecordBytes: MAX }, "t.bin"))
    expect(Exit.isFailure(cut) && JSON.stringify(cut.cause)).toContain("runs past the end of the part")
    const failing = Stream.concat(Stream.make(bytes), Stream.fail(new MalformedFile({ format: "zip", reason: "CRC-32 mismatch" })))
    const exit = await Effect.runPromiseExit(readRecords(failing, () => undefined, { maxRecordBytes: MAX }, "t.bin"))
    expect(Exit.isFailure(exit) && JSON.stringify(exit.cause)).toContain("CRC-32 mismatch")
  })

  it("readRecords gives the number of records read", async () => {
    const bytes = concat([record(0x83), record(0x9C, [1, 2, 3]), record(0x1), record(0x84)])
    expect(await Effect.runPromise(readRecords(Stream.make(bytes.subarray(0, 3), bytes.subarray(3)), () => undefined, { maxRecordBytes: MAX }, "t.bin"))).toBe(4)
    expect(await Effect.runPromise(readRecords(Stream.empty, () => undefined, { maxRecordBytes: MAX }, "t.bin"))).toBe(0)
  })

  it("a handler's stop ends the read without an error: no more bytes are pulled and the part's end isn't checked", async () => {
    const bytes = concat([record(0x83), record(0x167, [1, 2]), record(0x84)])
    const seen: Array<number> = []
    const stopAt = (r: Biff12Record) => {
      seen.push(r.type)
      return r.type === 0x167 ? "stop" as const : undefined
    }
    // a cut-off record and a stream error after the stop are never seen
    const rest = Stream.concat(Stream.make(new Uint8Array([0x85])), Stream.fail(new MalformedFile({ format: "zip", reason: "CRC-32 mismatch" })))
    const records = counter("records")
    const count = await Effect.runPromise(metered(records, readRecords(Stream.concat(Stream.make(bytes), rest), stopAt, { maxRecordBytes: MAX }, "t.bin")))
    expect(seen).toEqual([0x83, 0x167])
    expect(count).toBe(2)
    expect(records.total()).toBe(2)
    // the splitter hands the stop back, like an error, at the record that stopped
    const splitter = new RecordSplitter("t.bin", MAX)
    expect(splitter.feed(bytes, (r) => (r.type === 0x167 ? "stop" : undefined))).toBe("stop")
    expect(splitter.offset).toBe(record(0x83).length + record(0x167, [1, 2]).length)
  })
})

describe("fixed record layouts", () => {
  it("a record is checked once against its declared minimum: at the minimum it reads, below it is undefined", () => {
    const reader = new FixedReader()
    const data = new Uint8Array([...u32(0xFFFFFFFE), ...f64(-2.5), 0x34, 0x12, 0x7F])
    expect(reader.fixed(data, data.length + 1)).toBeUndefined()
    const f = reader.fixed(data, data.length)!
    expect([f.u32(0), f.i32(0), f.f64(4), f.u16(12), f.u8(14)]).toEqual([0xFFFFFFFE, -2, -2.5, 0x1234, 0x7F])
    expect(reader.fixed(new Uint8Array(0), 0)?.data.length).toBe(0)
  })

  it("reads a record's own bytes when it is a view inside a larger buffer", () => {
    const buffer = new Uint8Array([9, 9, 9, ...f64(1e300), ...u32(7), 9, 9])
    const f = new FixedReader().fixed(buffer.subarray(3, 15), 12)!
    expect([f.f64(0), f.u32(8)]).toEqual([1e300, 7])
  })
})

describe("biff12 strings", () => {
  it("reads an XLWideString and the offset after it", () => {
    const data = new Uint8Array([0xEE, ...wide("Sheet1"), 0xDD])
    expect(readWideString(data, 1, 100)).toEqual({ _tag: "Ok", value: "Sheet1", end: 1 + 4 + 12 })
    expect(readWideString(new Uint8Array(wide("")), 0, 100)).toEqual({ _tag: "Ok", value: "", end: 4 })
    expect(readWideString(new Uint8Array(wide("é😀")), 0, 100)).toEqual({ _tag: "Ok", value: "é😀", end: 4 + 6 })
  })

  it("past the data is Truncated, whatever the limit; else a count over the limit is TooLong with that count, before decoding", () => {
    const data = new Uint8Array(wide("abcdef"))
    expect(readWideString(data.subarray(0, 15), 0, 100)).toEqual({ _tag: "Truncated" })
    expect(readWideString(data, 2, 100)).toEqual({ _tag: "Truncated" })
    expect(readWideString(data, -1, 100)).toEqual({ _tag: "Truncated" })
    expect(readWideString(data, 0, 5)).toEqual({ _tag: "TooLong", count: 6 })
    expect(readWideString(data, 0, 6)).toEqual({ _tag: "Ok", value: "abcdef", end: 16 })
    // the real length, not the limit + 1
    expect(readWideString(data, 0, 2)).toEqual({ _tag: "TooLong", count: 6 })
    expect(readWideString(data, 0, 0)).toEqual({ _tag: "TooLong", count: 6 })
    // a huge count in a tiny record is a broken record (past the data), whatever the limit
    const huge = new Uint8Array(u32(0x7FFFFFFF))
    expect(readWideString(huge, 0, 1000)).toEqual({ _tag: "Truncated" })
    expect(readWideString(huge, 0, Infinity)).toEqual({ _tag: "Truncated" })
  })

  it("an XLNullableWideString of count 0xFFFFFFFF is no string", () => {
    expect(readNullableWideString(new Uint8Array(nullableWide(undefined)), 0, 100)).toEqual({ _tag: "Ok", value: undefined, end: 4 })
    expect(readNullableWideString(new Uint8Array(nullableWide("rId1")), 0, 100)).toEqual({ _tag: "Ok", value: "rId1", end: 12 })
    expect(readNullableWideString(new Uint8Array(3), 0, 100)).toEqual({ _tag: "Truncated" })
  })

  it("a RichStr's text follows its flags byte; runs and phonetic data after it are ignored", () => {
    const rich = new Uint8Array([0x03, ...wide("hello"), ...u32(1), 0, 0, 1, 0, 9, 9, 9])
    expect(readRichStr(rich, 0, 100)).toEqual({ _tag: "Ok", value: "hello", end: 15 })
    expect(readRichStr(new Uint8Array(0), 0, 100)).toEqual({ _tag: "Truncated" })
    expect(readRichStr(rich, 0, 4)).toEqual({ _tag: "TooLong", count: 5 })
  })
})

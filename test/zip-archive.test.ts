import { readdirSync, readFileSync, statSync } from "node:fs"
import { Effect, Exit, Stream } from "effect"
import * as Result from "effect/Result"
import { describe, expect, it } from "vitest"
import { WorkMeter } from "../src/internal/meter.ts"
import { type Archive, inflateBudget, type LocatedEntry, open } from "../src/internal/zip/archive.ts"
import { crc32, CRC32_START, crc32Final, crc32Update } from "../src/internal/zip/crc32.ts"
import { readDirectory, type ZipEntry } from "../src/internal/zip/directory.ts"
import type { Inflater, InflateRaw } from "../src/internal/zip/inflate.ts"
import { SLICE_BYTES } from "../src/internal/bytes.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { type Counter, counter } from "./helpers/meter.ts"
import { failureOf } from "./helpers/run.ts"
import { countYields } from "./helpers/yields.ts"
import { buildZip, concat, deflateRaw, readAll, testCrc32 } from "./helpers/zip.ts"

const limits = SpreadsheetLimits.defaults
const enc = (s: string) => new TextEncoder().encode(s)
/** The web inflater emits chunks of at most 16 KiB; limits are checked per chunk. */
const CHUNK = 64 * 1024

/** An opened archive whose reads all count on one work meter: `inflated` is the bytes inflated (or copied) over all its reads. */
interface Opened extends Archive {
  readonly inflated: Counter
}

const openOk = (bytes: Uint8Array, l = limits, inflate?: InflateRaw): Opened => {
  const opened = inflate === undefined ? open(bytes, l) : open(bytes, l, inflate)
  if (Result.isFailure(opened)) throw new Error(`open failed: ${opened.failure._tag} ${(opened.failure as { reason?: string }).reason}`)
  const archive = opened.success
  const inflated = counter("bytes")
  return { ...archive, read: (entry, l) => Stream.provideService(archive.read(entry, l), WorkMeter, inflated), inflated }
}

const readAllExit = (archive: Archive, name: string) =>
  Effect.runPromise(Effect.exit(readAll(archive, archive.find(name)!)))

const errorOf = async (archive: Archive, name: string) => failureOf(await readAllExit(archive, name))

const pseudoRandom = (n: number, seed = 1): Uint8Array => {
  const out = new Uint8Array(n)
  let x = seed
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0
    out[i] = x >>> 24
  }
  return out
}

describe("crc32", () => {
  it("matches the standard check value and an independent implementation", () => {
    expect(crc32(enc("123456789"))).toBe(0xCBF43926)
    const data = pseudoRandom(100_000)
    expect(crc32(data)).toBe(testCrc32(data))
    expect(crc32(new Uint8Array(0))).toBe(0)
    // every length mod 4, and running updates split anywhere
    for (let n = 1; n <= 33; n++) expect(crc32(data.subarray(0, n)), `n=${n}`).toBe(testCrc32(data.subarray(0, n)))
    for (const cut of [1, 2, 3, 5, 4097]) {
      expect(crc32Final(crc32Update(crc32Update(CRC32_START, data.subarray(0, cut)), data.subarray(cut)))).toBe(testCrc32(data))
    }
  })
})

describe("zip archive: reading entries", () => {
  it("reads stored and deflated entries back exactly; names are case-insensitive", async () => {
    const text = enc("<worksheet>" + "<row/>".repeat(10_000) + "</worksheet>")
    const random = pseudoRandom(200_000)
    const zip = await buildZip([
      { name: "xl/worksheets/sheet1.xml", data: text },
      { name: "[Content_Types].xml", data: enc("<Types/>"), method: "store" },
      { name: "xl/media/random.bin", data: random },
      { name: "empty.txt", data: new Uint8Array(0) }
    ])
    const archive = openOk(zip)
    expect(archive.entries.map((e) => e.name)).toEqual([
      "xl/worksheets/sheet1.xml",
      "[Content_Types].xml",
      "xl/media/random.bin",
      "empty.txt"
    ])
    expect(archive.find("/XL/Worksheets/Sheet1.XML")?.name).toBe("xl/worksheets/sheet1.xml")
    expect(archive.find("missing.xml")).toBeUndefined()
    for (const [name, data] of [["xl/worksheets/sheet1.xml", text], ["[content_types].xml", enc("<Types/>")], ["xl/media/random.bin", random], ["empty.txt", new Uint8Array(0)]] as const) {
      const exit = await readAllExit(archive, name)
      expect(Exit.isSuccess(exit) && Array.from(exit.value)).toEqual(Array.from(data))
    }
    // streams are re-runnable and come in bounded chunks
    const chunks = await Effect.runPromise(Stream.runCollect(archive.read(archive.find("xl/media/random.bin")!)))
    expect(Array.from(chunks).every((c) => c.length <= CHUNK)).toBe(true)
    expect(concat(Array.from(chunks))).toEqual(random)
  })

  it("reads every entry of the committed corpus workbooks, hidden-row fixtures and renamed copies included", async () => {
    // every zip in these directories, whatever its extension (renamed/ holds
    // workbooks saved under a misleading name)
    const dirs = ["test/corpus/generated", "test/corpus/hidden-rows", "test/corpus/hidden-rows/renamed"]
    let files = 0
    for (const dir of dirs) {
      for (const name of readdirSync(dir)) {
        if (!statSync(`${dir}/${name}`).isFile()) continue
        const bytes = new Uint8Array(readFileSync(`${dir}/${name}`))
        if (!(bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04)) continue
        const archive = openOk(bytes)
        for (const entry of archive.entries) {
          const exit = await Effect.runPromise(Effect.exit(readAll(archive, entry)))
          expect(Exit.isSuccess(exit), `${name}: ${entry.name}`).toBe(true)
        }
        // work done: exactly the declared sizes, once
        expect(archive.inflated.total()).toBe(archive.entries.reduce((n, e) => n + e.uncompressedSize, 0))
        files++
      }
    }
    expect(files).toBeGreaterThan(10)
  })

  it("reads the stored entries of the Python-written ZIP64 fixture", async () => {
    const archive = openOk(new Uint8Array(readFileSync("test/fixtures/zip/zip64.zip")))
    const exit = await readAllExit(archive, "xl/workbook.xml")
    expect(Exit.isSuccess(exit) && new TextDecoder().decode(exit.value)).toBe("<workbook/>")
  })

  it("reads sizes and offsets from a ZIP64 extra field", async () => {
    const zip = await buildZip([
      { name: "a.xml", data: enc("first") },
      { name: "b.xml", data: enc("second entry"), zip64: true }
    ])
    const archive = openOk(zip)
    const b = archive.find("b.xml")!
    expect([b.uncompressedSize, b.localOffset > 0]).toEqual([12, true])
    const exit = await readAllExit(archive, "b.xml")
    expect(Exit.isSuccess(exit) && new TextDecoder().decode(exit.value)).toBe("second entry")
  })

  it("the inflate implementation is replaceable; a stopped read cancels it", async () => {
    let cancelled = 0
    let pulls = 0
    const fake: InflateRaw = () => {
      const inflater: Inflater = {
        next: async () => {
          pulls++
          return pulls > 1000 ? undefined : new Uint8Array(1024)
        },
        cancel: () => {
          cancelled++
        }
      }
      return inflater
    }
    const zip = await buildZip([{ name: "big.bin", data: new Uint8Array(1000 * 1024) }])
    const archive = openOk(zip, limits, fake)
    const first = await Effect.runPromise(Stream.runCollect(Stream.take(archive.read(archive.find("big.bin")!), 2)))
    expect(Array.from(first).length).toBe(2)
    expect(pulls).toBe(2)
    expect(cancelled).toBeGreaterThanOrEqual(1)
  })
})

describe("zip archive: bounded slices", () => {
  const SIZE = 1024 * 1024 + 123
  const data = pseudoRandom(SIZE, 7)
  /** An inflater that hands the whole entry over in one chunk. */
  const onePiece: InflateRaw = () => {
    let given = false
    return {
      next: async () => {
        if (given) return undefined
        given = true
        return data.slice()
      },
      cancel: () => {}
    }
  }

  const slicesOf = async (archive: Archive, name: string) => {
    const pulls = await Effect.runPromise(Stream.runCollect(Stream.chunks(archive.read(archive.find(name)!))))
    return Array.from(pulls)
  }

  for (const [label, method, inflate] of [
    ["stored", "store", undefined],
    ["deflated", "deflate", undefined],
    ["deflated in one piece", "deflate", onePiece]
  ] as const) {
    it(`a ${label} entry is read one slice of at most SLICE_BYTES per pull, each after a yield`, async () => {
      const zip = await buildZip([{ name: "big.bin", data, method }])
      const archive = openOk(zip, limits, inflate)
      const run = await countYields(() => slicesOf(archive, "big.bin"))
      const pulls = run.value
      expect(pulls.every((pull) => pull.length === 1)).toBe(true)
      const slices = pulls.map((pull) => pull[0]!)
      expect(slices.every((slice) => slice.length > 0 && slice.length <= SLICE_BYTES)).toBe(true)
      expect(Buffer.from(concat(slices)).equals(data)).toBe(true)
      expect(slices.length).toBeGreaterThanOrEqual(Math.ceil(SIZE / SLICE_BYTES))
      expect(run.yields).toBeGreaterThanOrEqual(slices.length / 2)
    })
  }
})

describe("zip64 extra field bounds", () => {
  /** A one-entry archive whose directory entry uses a ZIP64 extra field; returns the bytes and where the field is. */
  const zip64Archive = async () => {
    const bytes = await buildZip([{ name: "a.xml", data: enc("hello"), zip64: true }])
    const view = new DataView(bytes.buffer)
    const cd = view.getUint32(bytes.length - 22 + 16, true)
    const extra = cd + 46 + "a.xml".length // header id, size, then usize, csize, offset
    return { bytes, view, extra }
  }
  const reasonOf = (result: ReturnType<typeof readDirectory>) =>
    Result.isSuccess(result) ? "ok" : (result.failure as { reason?: string }).reason ?? result.failure._tag

  it("a ZIP64 field too short for the values it must hold is MalformedFile", async () => {
    const { bytes, view, extra } = await zip64Archive()
    view.setUint16(extra + 2, 8, true) // room for the uncompressed size only
    expect(reasonOf(readDirectory(bytes, limits))).toBe("bad ZIP64 extra field")
  })

  it("an extra field longer than the extra area is ignored, so the ZIP64 values are missing", async () => {
    const { bytes, view, extra } = await zip64Archive()
    view.setUint16(extra + 2, 100, true)
    expect(reasonOf(readDirectory(bytes, limits))).toBe("ZIP64 sizes without a ZIP64 extra field")
  })

  it("a ZIP64 offset of 2^60 can't be a buffer offset", async () => {
    const { bytes, view, extra } = await zip64Archive()
    view.setBigUint64(extra + 4 + 16, 1n << 60n, true)
    expect(reasonOf(readDirectory(bytes, limits))).toBe("bad ZIP64 extra field")
  })

  it("a ZIP64 compressed size of 2^53-1 is read exactly and fails as out of bounds", async () => {
    const { bytes, view, extra } = await zip64Archive()
    view.setBigUint64(extra + 4 + 8, (1n << 53n) - 1n, true)
    const { entries } = Result.getOrThrow(readDirectory(bytes, limits))
    expect(entries[0]!.compressedSize).toBe(Number.MAX_SAFE_INTEGER)
    expect(Result.merge(open(bytes, limits))).toMatchObject({ _tag: "MalformedFile", reason: "entry data out of bounds" })
  })
})

describe("zip archive: hostile input", () => {
  const MiB = 1024 * 1024
  const FLOOR = 64 * MiB

  it("the archive-wide budget is max(64 MiB, 200 x archive size), capped at maxDecompressedBytes", () => {
    expect(inflateBudget(10_000, limits)).toBe(FLOOR)
    expect(inflateBudget(1 * MiB, limits)).toBe(200 * MiB)
    expect(inflateBudget(100 * MiB, limits)).toBe(1024 * MiB)
  })

  it("a single-entry zip bomb stops just past the 64 MiB floor", async () => {
    const zeros = new Uint8Array(96 * MiB)
    const zip = await buildZip([{ name: "bomb.xml", data: zeros }])
    expect(zip.length).toBeLessThan(200_000)
    const archive = openOk(zip)
    const error = await errorOf(archive, "bomb.xml")
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxInflateRatio", max: 200 })
    // work done: the budget plus at most one chunk, not 96 MiB
    expect(archive.inflated.total()).toBeLessThanOrEqual(FLOOR + CHUNK)
  }, 60_000)

  it("a bomb that would inflate forever stops at the budget (fake inflater, declared 1 GiB)", async () => {
    const chunk = new Uint8Array(MiB)
    const endless: InflateRaw = () => ({ next: async () => chunk, cancel: () => {} })
    const zip = await buildZip([{ name: "endless.xml", data: enc("x"), uncompressedSize: 1024 * MiB }])
    const archive = openOk(zip, limits, endless)
    expect(await errorOf(archive, "endless.xml")).toMatchObject({ _tag: "LimitExceeded", limit: "maxInflateRatio" })
    expect(archive.inflated.total()).toBeLessThanOrEqual(FLOOR + MiB)
  })

  const zeroEntries = async (count: number) => {
    // one deflated MiB of zeros, reused for every entry (each well under the
    // old per-entry 1 MiB grace, so only an archive-wide budget sees them)
    const zeros = new Uint8Array(MiB)
    const deflated = await deflateRaw(zeros)
    const crc = testCrc32(zeros)
    return buildZip(Array.from({ length: count }, (_, i) => ({
      name: `zeros${i}.xml`, data: deflated, method: "store" as const, methodCode: 8, uncompressedSize: MiB, crc32: crc
    })))
  }
  const readEntries = async (archive: Archive) => {
    let read = 0
    for (const entry of archive.entries) {
      const error = failureOf(await Effect.runPromise(Effect.exit(readAll(archive, entry))))
      if (error !== undefined) return { read, error }
      read++
    }
    return { read, error: undefined }
  }

  it("many small 1 MiB-zero entries are stopped by the archive-wide budget", async () => {
    const zip = await zeroEntries(100)
    expect(zip.length).toBeLessThan(200 * 1024)
    const archive = openOk(zip)
    const { read, error } = await readEntries(archive)
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxInflateRatio", max: 200 })
    expect(read).toBe(64) // the 65th MiB crosses the floor
    expect(archive.inflated.total()).toBeLessThanOrEqual(FLOOR + CHUNK)
  }, 60_000)

  it("50 such entries (50 MiB) fit within the 64 MiB floor; a lower floor stops them early", async () => {
    const zip = await zeroEntries(50)
    expect((await readEntries(openOk(zip))).error).toBeUndefined()
    // with a 1 MiB floor the ratio part decides: 200 x the archive size
    const tightLimits = { ...limits, inflateFloorBytes: MiB }
    const budget = inflateBudget(zip.length, tightLimits)
    expect(budget).toBe(200 * zip.length)
    const tight = openOk(zip, tightLimits)
    const { read, error } = await readEntries(tight)
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxInflateRatio" })
    expect(read).toBe(Math.floor(budget / MiB))
    expect(read).toBeLessThan(50)
    expect(tight.inflated.total()).toBeLessThanOrEqual(budget + CHUNK)
  }, 60_000)

  it("reading an entry again doesn't spend the budget again", async () => {
    const zip = await zeroEntries(1)
    const archive = openOk(zip, { ...limits, inflateFloorBytes: 2 * MiB })
    for (let k = 0; k < 5; k++) expect((await readEntries(archive)).error).toBeUndefined()
    expect(archive.inflated.total()).toBe(5 * MiB)
  })

  it("after one entry goes over the budget, entries already read still read; new bytes still fail", async () => {
    const zip = await zeroEntries(3)
    const archive = openOk(zip, { ...limits, inflateFloorBytes: 2 * MiB })
    const [a, b, c] = archive.entries
    const readExit = (entry: LocatedEntry | undefined, l?: SpreadsheetLimits.Limits) => Effect.runPromise(Effect.exit(Stream.runCollect(archive.read(entry!, l))))
    expect(Exit.isSuccess(await readExit(a))).toBe(true)
    expect(Exit.isSuccess(await readExit(b))).toBe(true)
    const over = { _tag: "LimitExceeded", limit: "maxInflateRatio", max: 200 }
    expect(failureOf(await readExit(c))).toMatchObject(over)
    // a and b spend nothing new: they still read
    expect(Exit.isSuccess(await readExit(a))).toBe(true)
    expect(Exit.isSuccess(await readExit(b))).toBe(true)
    // c fails again once it passes what it already inflated
    expect(failureOf(await readExit(c))).toMatchObject(over)
    // limits given to one read set its budget against the same total
    expect(Exit.isSuccess(await readExit(c, { ...limits, inflateFloorBytes: 3 * MiB }))).toBe(true)
    expect(failureOf(await readExit(a, { ...limits, inflateFloorBytes: 0, maxInflateRatio: 1 }))).toBeUndefined()
    expect(archive.inflated.total()).toBeLessThanOrEqual(6 * MiB + 2 * CHUNK)
  })

  it("the chunk that goes over the budget isn't charged, so reading the entry again fails again", async () => {
    // stored: read in three 64 KiB slices; the third one crosses 150 KiB
    const zip = await buildZip([{ name: "s.bin", data: pseudoRandom(3 * 64 * 1024), method: "store" }])
    const archive = openOk(zip, { ...limits, maxDecompressedBytes: 150 * 1024 })
    const over = { _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: 150 * 1024, actual: 3 * 64 * 1024 }
    expect(await errorOf(archive, "s.bin")).toMatchObject(over)
    expect(await errorOf(archive, "s.bin")).toMatchObject(over)
  })

  it("the committed 666:1 blank-row workbook passes with the default limits", async () => {
    const bytes = new Uint8Array(readFileSync("test/corpus/generated/many_empty_rows_implicit.xlsx"))
    const archive = openOk(bytes)
    expect((await readEntries(archive)).error).toBeUndefined()
  })

  it("the total budget covers all entries together", async () => {
    // compressible but well under the ratio limit
    const block = concat(Array.from({ length: 8 }, (_, i) => pseudoRandom(128 * 1024, i + 1)))
    const zip = await buildZip(Array.from({ length: 10 }, (_, i) => ({ name: `part${i}.bin`, data: block })))
    const archive = openOk(zip, { ...limits, maxDecompressedBytes: 3 * 1024 * 1024 })
    let error: unknown
    for (const entry of archive.entries) {
      error = failureOf(await Effect.runPromise(Effect.exit(readAll(archive, entry))))
      if (error !== undefined) break
    }
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes" })
    expect(archive.inflated.total()).toBeLessThanOrEqual(3 * 1024 * 1024 + CHUNK)
  })

  it("an entry that inflates past its declared size stops there", async () => {
    const data = new Uint8Array(4 * 1024 * 1024).fill(65)
    const zip = await buildZip([{ name: "lie.xml", data, uncompressedSize: 1000 }])
    const archive = openOk(zip)
    const error = await errorOf(archive, "lie.xml")
    expect(error).toMatchObject({ _tag: "MalformedFile", part: "lie.xml" })
    expect(String((error as { reason: string }).reason)).toMatch(/past its declared size/)
    expect(archive.inflated.total()).toBeLessThanOrEqual(1000 + CHUNK)
  })

  it("an entry shorter than declared, or with a wrong CRC, is MalformedFile", async () => {
    const zip = await buildZip([
      { name: "short.xml", data: enc("abc"), uncompressedSize: 10 },
      { name: "crc.xml", data: enc("abc"), crc32: 12345 }
    ])
    const archive = openOk(zip)
    expect(await errorOf(archive, "short.xml")).toMatchObject({ _tag: "MalformedFile", part: "short.xml" })
    const crc = await errorOf(archive, "crc.xml")
    expect(crc).toMatchObject({ _tag: "MalformedFile", part: "crc.xml", reason: "CRC-32 mismatch" })
  })

  it("truncated, garbage and trailing-junk deflate data are MalformedFile", async () => {
    const good = await deflateRaw(enc("hello world ".repeat(1000)))
    const data = enc("hello world ".repeat(1000))
    const zip = await buildZip([
      { name: "trailing.xml", data, trailing: enc("junk") },
      { name: "garbage.xml", data: new Uint8Array(16), method: "store", methodCode: 8, compressedSize: 16, uncompressedSize: 16 }
    ])
    const archive = openOk(zip)
    for (const name of ["trailing.xml", "garbage.xml"]) {
      const error = await errorOf(archive, name)
      expect(error, name).toMatchObject({ _tag: "MalformedFile", part: name })
      expect(String((error as { reason: string }).reason), name).toMatch(/invalid deflate data/)
    }
    // truncated: cut the compressed stream but keep the declared sizes
    const truncated = await buildZip([{ name: "cut.xml", data: good.subarray(0, good.length - 10), method: "store", methodCode: 8, uncompressedSize: data.length, crc32: testCrc32(data) }])
    const cut = openOk(truncated)
    expect(await errorOf(cut, "cut.xml")).toMatchObject({ _tag: "MalformedFile", part: "cut.xml" })
  })

  it("encrypted entries are Encrypted, unknown methods UnsupportedFormat", async () => {
    const zip = await buildZip([
      { name: "secret.xml", data: enc("x"), flags: 0x801 },
      { name: "bzip.xml", data: enc("x"), method: "store", methodCode: 12 }
    ])
    const archive = openOk(zip)
    expect(await errorOf(archive, "secret.xml")).toMatchObject({ _tag: "Encrypted", format: "zip" })
    expect(await errorOf(archive, "bzip.xml")).toMatchObject({ _tag: "UnsupportedFormat", format: "zip" })
  })

  it("open classifies how each entry reads; a read fails with the first reason that applies", async () => {
    const zip = await buildZip([
      { name: "plain.xml", data: enc("stored"), method: "store" },
      { name: "packed.xml", data: enc("deflated deflated deflated") },
      { name: "secret.xml", data: enc("x"), flags: 0x801 },
      // encrypted with an unknown method: encryption is reported
      { name: "secret-bzip.xml", data: enc("x"), method: "store", methodCode: 12, flags: 0x801 },
      // encrypted and stored with sizes that differ: encryption is reported
      { name: "secret-short.xml", data: enc("xyz"), method: "store", uncompressedSize: 9, flags: 0x801 },
      { name: "bzip.xml", data: enc("x"), method: "store", methodCode: 12 },
      { name: "short.xml", data: enc("xyz"), method: "store", uncompressedSize: 9 }
    ])
    const archive = openOk(zip)
    // find returns the entry itself, not a copy: the inflate budget is kept per entry object
    for (const entry of archive.entries) {
      expect(archive.find(entry.name)).toBe(entry)
      expect(archive.find(entry.name.toUpperCase())).toBe(entry)
    }
    expect(Object.fromEntries(archive.entries.map((e) => [e.name, e.kind]))).toEqual({
      "plain.xml": "stored",
      "packed.xml": "deflate",
      "secret.xml": "encrypted",
      "secret-bzip.xml": "encrypted",
      "secret-short.xml": "encrypted",
      "bzip.xml": { unsupported: 12 },
      "short.xml": "bad-stored-size"
    })
    // each entry's data range is its bytes after the local header
    const plain = archive.find("plain.xml")!
    expect(new TextDecoder().decode(zip.subarray(plain.data.start, plain.data.end))).toBe("stored")
    expect(plain.data.start).toBe(plain.localOffset + 30 + "plain.xml".length)
    // a local extra field (here the 0x5455 timestamp field, absent from the directory) comes before the data
    const stamped = openOk(await buildZip([{ name: "t.xml", data: enc("timed"), method: "store", localExtra: new Uint8Array([0x55, 0x54, 5, 0, 1, 0, 0, 0, 0]) }]))
    const timed = stamped.find("t.xml")!
    expect(timed.data).toEqual({ start: 30 + "t.xml".length + 9, end: 30 + "t.xml".length + 9 + 5 })
    expect(new TextDecoder().decode(await Effect.runPromise(readAll(stamped, timed)))).toBe("timed")
    const packed = archive.find("PACKED.xml")!
    expect(packed.data.end - packed.data.start).toBe(packed.compressedSize)
    expect(new TextDecoder().decode(await Effect.runPromise(readAll(archive, plain)))).toBe("stored")
    expect(new TextDecoder().decode(await Effect.runPromise(readAll(archive, packed)))).toBe("deflated deflated deflated")
    expect(await errorOf(archive, "secret-bzip.xml")).toMatchObject({ _tag: "Encrypted", format: "zip" })
    expect(await errorOf(archive, "secret-short.xml")).toMatchObject({ _tag: "Encrypted", format: "zip" })
    expect(await errorOf(archive, "bzip.xml")).toMatchObject({ _tag: "UnsupportedFormat", format: "zip", reason: "compression method 12 in bzip.xml" })
    expect(await errorOf(archive, "short.xml")).toMatchObject({
      _tag: "MalformedFile",
      format: "zip",
      reason: "stored entry with different compressed and uncompressed sizes",
      part: "short.xml"
    })
    // only entries an archive has opened can be read: a bare directory entry has no checked data range
    const bare = (entry: ZipEntry) =>
      // @ts-expect-error a directory entry isn't a located entry
      archive.read(entry)
    expect(typeof bare).toBe("function")
  })

  it("layout problems fail at open: overlap, duplicate names, bad local headers, data past the end", async () => {
    const two = await buildZip([{ name: "a.xml", data: enc("aaaa"), method: "store" }, { name: "b.xml", data: enc("bbbb"), method: "store" }])
    const eocd = two.length - 22
    const view = new DataView(two.buffer)
    const cd = view.getUint32(eocd + 16, true)
    // a's compressed size grown into b's header
    const overlap = two.slice()
    new DataView(overlap.buffer).setUint32(cd + 20, 20, true)
    expect(Result.merge(open(overlap, limits))).toMatchObject({ _tag: "MalformedFile", reason: "entries overlap" })
    // duplicate names, compared case-insensitively
    const dup = await buildZip([{ name: "A.xml", data: enc("1") }, { name: "a.XML", data: enc("2") }])
    expect(Result.merge(open(dup, limits))).toMatchObject({ _tag: "MalformedFile", reason: "duplicate entry name" })
    // names are checked before the local headers: with both wrong, the duplicate is reported
    const dupBadSig = dup.slice()
    new DataView(dupBadSig.buffer).setUint32(0, 0, true)
    expect(Result.merge(open(dupBadSig, limits))).toMatchObject({ _tag: "MalformedFile", reason: "duplicate entry name" })
    // a local header with the wrong signature
    const badSig = two.slice()
    new DataView(badSig.buffer).setUint32(0, 0, true)
    expect(Result.merge(open(badSig, limits))).toMatchObject({ _tag: "MalformedFile", reason: "bad local header signature" })
    // a local name that differs from the directory's
    const badName = two.slice()
    badName[30] = "z".charCodeAt(0)
    expect(Result.merge(open(badName, limits))).toMatchObject({ _tag: "MalformedFile", reason: "local header name differs from the directory" })
    // names compare decoded: a Latin-1 local name ("é" as E9) is not the directory's UTF-8 "é" (C3 A9),
    // though its bytes are the directory name's character codes
    const latin1 = await buildZip([{ name: "\u00e9\u00e9", data: new Uint8Array(0), method: "store" }])
    const lv = new DataView(latin1.buffer)
    lv.setUint16(26, 2, true) // the local name is 2 bytes; the 2 left over become a gap before the directory
    latin1.set([0xE9, 0xE9], 30)
    expect(Result.merge(open(latin1, limits))).toMatchObject({ _tag: "MalformedFile", reason: "local header name differs from the directory" })
    // and the same names in UTF-8 match
    expect(Result.isSuccess(open(await buildZip([{ name: "\u00e9\u00e9", data: enc("x") }]), limits))).toBe(true)
    // a directory offset past the end of the data
    const farOffset = two.slice()
    new DataView(farOffset.buffer).setUint32(cd + 46 + 5 + 42, 1_000_000, true)
    expect(Result.merge(open(farOffset, limits))).toMatchObject({ _tag: "MalformedFile", reason: "local header out of bounds" })
    // a local header whose method differs from the directory's
    const badMethod = two.slice()
    new DataView(badMethod.buffer).setUint16(8, 8, true)
    expect(Result.merge(open(badMethod, limits))).toMatchObject({ _tag: "MalformedFile", reason: "local header compression method differs from the directory" })
    // b's data grown into the central directory (still inside the buffer)
    const intoDirectory = two.slice()
    new DataView(intoDirectory.buffer).setUint32(cd + 46 + 5 + 20, 10, true)
    expect(Result.merge(open(intoDirectory, limits))).toMatchObject({ _tag: "MalformedFile", reason: "entry data overlaps the central directory" })
    // compressed size past the buffer
    const farData = two.slice()
    new DataView(farData.buffer).setUint32(cd + 46 + 5 + 20, 1_000_000, true)
    expect(Result.merge(open(farData, limits))).toMatchObject({ _tag: "MalformedFile", reason: "entry data out of bounds" })
  })
})

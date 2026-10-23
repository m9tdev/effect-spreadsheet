import { readdirSync, readFileSync } from "node:fs"
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { type Archive, open, RATIO_GRACE_BYTES } from "../src/internal/zip/archive.ts"
import { crc32 } from "../src/internal/zip/crc32.ts"
import type { Inflater, InflateRaw } from "../src/internal/zip/inflate.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { failureOf } from "./helpers/run.ts"
import { buildZip, concat, deflateRaw, testCrc32 } from "./helpers/zip.ts"

const limits = SpreadsheetLimits.defaults
const enc = (s: string) => new TextEncoder().encode(s)
/** The web inflater emits chunks of at most 16 KiB; limits are checked per chunk. */
const CHUNK = 64 * 1024

const openOk = (bytes: Uint8Array, l = limits, inflate?: InflateRaw): Archive => {
  const archive = inflate === undefined ? open(bytes, l) : open(bytes, l, inflate)
  if ("_tag" in archive) throw new Error(`open failed: ${archive._tag} ${(archive as { reason?: string }).reason}`)
  return archive
}

const readAllExit = (archive: Archive, name: string) =>
  Effect.runPromise(Effect.exit(archive.readAll(archive.find(name)!)))

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

  it("reads every entry of the committed corpus workbooks (CRCs and sizes check out)", async () => {
    const dirs = ["test/corpus/generated", "test/corpus/hidden-rows"]
    let files = 0
    for (const dir of dirs) {
      for (const name of readdirSync(dir).filter((n) => /\.(xlsx|xlsb|ods)$/.test(n))) {
        const bytes = new Uint8Array(readFileSync(`${dir}/${name}`))
        // some fixtures are deliberately highly compressible
        const archive = openOk(bytes, { ...limits, maxCompressionRatio: 10_000 })
        for (const entry of archive.entries) {
          const exit = await Effect.runPromise(Effect.exit(archive.readAll(entry)))
          expect(Exit.isSuccess(exit), `${name}: ${entry.name}`).toBe(true)
        }
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

describe("zip archive: hostile input", () => {
  it("a zip bomb stops at the ratio limit after inflating about the grace amount", async () => {
    const zeros = new Uint8Array(64 * 1024 * 1024)
    const zip = await buildZip([{ name: "bomb.xml", data: zeros }])
    expect(zip.length).toBeLessThan(200_000)
    const archive = openOk(zip)
    const error = await errorOf(archive, "bomb.xml")
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxCompressionRatio", max: 100 })
    // work done: up to the ratio (or the grace amount) plus one chunk, not 64 MiB
    const compressed = archive.find("bomb.xml")!.compressedSize
    expect(archive.inflatedBytes()).toBeLessThanOrEqual(Math.max(RATIO_GRACE_BYTES, 100 * compressed) + CHUNK)
    expect(archive.inflatedBytes()).toBeLessThan(zeros.length / 8)
  }, 30_000)

  it("the total budget covers all entries together", async () => {
    // compressible but well under the ratio limit
    const block = concat(Array.from({ length: 8 }, (_, i) => pseudoRandom(128 * 1024, i + 1)))
    const zip = await buildZip(Array.from({ length: 10 }, (_, i) => ({ name: `part${i}.bin`, data: block })))
    const archive = openOk(zip, { ...limits, maxDecompressedBytes: 3 * 1024 * 1024 })
    let error: unknown
    for (const entry of archive.entries) {
      error = failureOf(await Effect.runPromise(Effect.exit(archive.readAll(entry))))
      if (error !== undefined) break
    }
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes" })
    expect(archive.inflatedBytes()).toBeLessThanOrEqual(3 * 1024 * 1024 + CHUNK)
  })

  it("an entry that inflates past its declared size stops there", async () => {
    const data = new Uint8Array(4 * 1024 * 1024).fill(65)
    const zip = await buildZip([{ name: "lie.xml", data, uncompressedSize: 1000 }])
    const archive = openOk(zip)
    const error = await errorOf(archive, "lie.xml")
    expect(error).toMatchObject({ _tag: "MalformedFile", part: "lie.xml" })
    expect(String((error as { reason: string }).reason)).toMatch(/past its declared size/)
    expect(archive.inflatedBytes()).toBeLessThanOrEqual(1000 + CHUNK)
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

  it("layout problems fail at open: overlap, duplicate names, bad local headers, data past the end", async () => {
    const two = await buildZip([{ name: "a.xml", data: enc("aaaa"), method: "store" }, { name: "b.xml", data: enc("bbbb"), method: "store" }])
    const eocd = two.length - 22
    const view = new DataView(two.buffer)
    const cd = view.getUint32(eocd + 16, true)
    // a's compressed size grown into b's header
    const overlap = two.slice()
    new DataView(overlap.buffer).setUint32(cd + 20, 20, true)
    expect(open(overlap, limits)).toMatchObject({ _tag: "MalformedFile", reason: "entries overlap" })
    // duplicate names, compared case-insensitively
    const dup = await buildZip([{ name: "A.xml", data: enc("1") }, { name: "a.XML", data: enc("2") }])
    expect(open(dup, limits)).toMatchObject({ _tag: "MalformedFile", reason: "duplicate entry name" })
    // a local header with the wrong signature
    const badSig = two.slice()
    new DataView(badSig.buffer).setUint32(0, 0, true)
    expect(open(badSig, limits)).toMatchObject({ _tag: "MalformedFile", reason: "bad local header signature" })
    // a local name that differs from the directory's
    const badName = two.slice()
    badName[30] = "z".charCodeAt(0)
    expect(open(badName, limits)).toMatchObject({ _tag: "MalformedFile", reason: "local header name differs from the directory" })
    // a directory offset past the end of the data
    const farOffset = two.slice()
    new DataView(farOffset.buffer).setUint32(cd + 46 + 5 + 42, 1_000_000, true)
    expect(open(farOffset, limits)).toMatchObject({ _tag: "MalformedFile", reason: "local header out of bounds" })
    // compressed size past the buffer
    const farData = two.slice()
    new DataView(farData.buffer).setUint32(cd + 46 + 5 + 20, 1_000_000, true)
    expect(open(farData, limits)).toMatchObject({ _tag: "MalformedFile", reason: "entry data out of bounds" })
  })
})

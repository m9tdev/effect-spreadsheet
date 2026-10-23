import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { readDirectory, type ZipEntry } from "../src/internal/zip/directory.ts"
import { detect } from "../src/internal/sniff.ts"
import { makeZip } from "./helpers/zip.ts"

const limits = SpreadsheetLimits.defaults
const names = (result: ReturnType<typeof readDirectory>) =>
  Array.isArray(result) ? (result as ReadonlyArray<ZipEntry>).map((e) => e.name) : `${(result as { _tag: string })._tag}`
const reason = (result: ReturnType<typeof readDirectory>) =>
  Array.isArray(result) ? "ok" : (result as { reason?: string; _tag: string }).reason ?? (result as { _tag: string })._tag

// Written by Python's zipfile with the ZIP64 records forced on (see make-zip64.py).
const ZIP64 = new Uint8Array(readFileSync(new URL("./fixtures/zip/zip64.zip", import.meta.url)))
const EXPECTED = ["[Content_Types].xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml"]

const eocdOf = (bytes: Uint8Array): number => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let i = bytes.length - 22; i >= 0; i--) if (view.getUint32(i, true) === 0x06054b50) return i
  throw new Error("no EOCD")
}

/** The fixture with the classic EOCD fields set to "see ZIP64". */
const zip64Patched = (): Uint8Array => {
  const bytes = ZIP64.slice()
  const view = new DataView(bytes.buffer)
  const eocd = eocdOf(bytes)
  view.setUint16(eocd + 8, 0xFFFF, true)
  view.setUint16(eocd + 10, 0xFFFF, true)
  view.setUint32(eocd + 12, 0xFFFFFFFF, true)
  view.setUint32(eocd + 16, 0xFFFFFFFF, true)
  return bytes
}

const withComment = (zip: Uint8Array, commentLength: number, trailing = 0): Uint8Array => {
  const out = new Uint8Array(zip.length + commentLength + trailing)
  out.set(zip)
  new DataView(out.buffer).setUint16(zip.length - 2, Math.min(commentLength, 0xFFFF), true)
  return out
}

describe("zip central directory", () => {
  it("reads Python's ZIP64 archive, unpatched and with the EOCD saying 0xFFFF/0xFFFFFFFF", () => {
    expect(names(readDirectory(ZIP64, limits))).toEqual(EXPECTED)
    const patched = zip64Patched()
    expect(names(readDirectory(patched, limits))).toEqual(EXPECTED)
    expect(detect(patched, limits)).toBe("xlsx")
    // the ZIP64 entry count is checked against maxZipEntries
    expect(names(readDirectory(patched, { ...limits, maxZipEntries: 2 }))).toBe("LimitExceeded")
  })

  it("ZIP64 record fields are read at the right offsets", () => {
    const patched = zip64Patched()
    const view = new DataView(patched.buffer)
    const locator = eocdOf(patched) - 20
    const at = Number(view.getBigUint64(locator + 8, true))
    // wrong ZIP64 central-directory size: entries run out of bounds
    const badSize = patched.slice()
    new DataView(badSize.buffer).setBigUint64(at + 40, 10n, true)
    expect(reason(readDirectory(badSize, limits))).toMatch(/out of bounds/)
    // wrong ZIP64 central-directory offset
    const badStart = patched.slice()
    new DataView(badStart.buffer).setBigUint64(at + 48, 5n, true)
    expect(reason(readDirectory(badStart, limits))).toMatch(/signature|out of bounds/)
    // ZIP64 count above 2^53 can't be real
    const hugeCount = patched.slice()
    new DataView(hugeCount.buffer).setBigUint64(at + 32, 1n << 60n, true)
    expect(reason(readDirectory(hugeCount, limits))).toMatch(/entry count/)
  })

  it("ZIP64 locator problems are MalformedFile", () => {
    const patched = zip64Patched()
    const locator = eocdOf(patched) - 20
    // locator points so close to itself that the 56-byte record would overlap it
    const overlapping = patched.slice()
    new DataView(overlapping.buffer).setBigUint64(locator + 8, BigInt(locator - 55), true)
    expect(reason(readDirectory(overlapping, limits))).toBe("ZIP64 end of central directory out of bounds")
    // exactly fitting (at + 56 === locator) passes the bound and reaches the signature check
    const view = new DataView(patched.buffer)
    const at = Number(view.getBigUint64(locator + 8, true))
    expect(at + 56).toBe(locator)
    // missing locator
    const noLocator = patched.slice()
    new DataView(noLocator.buffer).setUint32(locator, 0, true)
    expect(reason(readDirectory(noLocator, limits))).toBe("ZIP64 locator missing")
    // bad ZIP64 EOCD signature
    const badSig = patched.slice()
    new DataView(badSig.buffer).setUint32(at, 0, true)
    expect(reason(readDirectory(badSig, limits))).toBe("bad ZIP64 end of central directory")
  })

  it("finds the EOCD behind a 65535-byte comment, not behind 65536 bytes", () => {
    const zip = makeZip([{ name: "xl/workbook.xml" }])
    expect(names(readDirectory(withComment(zip, 65535), limits))).toEqual(["xl/workbook.xml"])
    // 65536 bytes after the EOCD record: past what a comment can be
    expect(reason(readDirectory(withComment(zip, 65535, 1), limits))).toMatch(/not found/)
  })

  it("a central directory running past the EOCD is MalformedFile, even if it fits in the buffer", () => {
    const zip = withComment(makeZip([{ name: "a" }, { name: "b" }]), 200)
    const eocd = eocdOf(zip)
    const view = new DataView(zip.buffer)
    const start = view.getUint32(eocd + 16, true)
    const size = view.getUint32(eocd + 12, true)
    expect(start + size).toBe(eocd)
    view.setUint32(eocd + 12, size + 100, true) // still < zip.length thanks to the comment
    expect(start + size + 100).toBeLessThan(zip.length)
    expect(reason(readDirectory(zip, limits))).toBe("central directory out of bounds")
  })

  it("with ZIP64, the central directory must end before the ZIP64 record", () => {
    const patched = zip64Patched()
    const view = new DataView(patched.buffer)
    const locator = eocdOf(patched) - 20
    const at = Number(view.getBigUint64(locator + 8, true))
    const start = Number(view.getBigUint64(at + 48, true))
    const size = Number(view.getBigUint64(at + 40, true))
    expect(start + size).toBe(at)
    // 10 bytes more still ends before the classic EOCD, but runs into the ZIP64 record
    view.setBigUint64(at + 40, BigInt(size + 10), true)
    expect(start + size + 10).toBeLessThan(eocdOf(patched))
    expect(reason(readDirectory(patched, limits))).toBe("central directory out of bounds")
  })

  it("an entry whose name runs past the end of the central directory is MalformedFile", () => {
    const zip = makeZip([{ name: "first" }, { name: "second" }])
    const eocd = eocdOf(zip)
    const view = new DataView(zip.buffer)
    // shrink the directory by 1 byte: the last entry's header fits, its name doesn't
    view.setUint32(eocd + 12, view.getUint32(eocd + 12, true) - 1, true)
    expect(reason(readDirectory(zip, limits))).toBe("central directory entry out of bounds")
  })

  it("an EOCD at offset 0 asking for ZIP64 is a typed error, not a RangeError", async () => {
    const bytes = new Uint8Array(22)
    const view = new DataView(bytes.buffer)
    view.setUint32(0, 0x06054b50, true)
    view.setUint16(8, 0xFFFF, true)
    view.setUint16(10, 0xFFFF, true)
    expect(reason(readDirectory(bytes, limits))).toBe("ZIP64 locator missing")
    expect(detect(bytes, limits)).toMatchObject({ _tag: "MalformedFile", format: "zip" })
  })

  it("an archive with both workbook parts is xlsb (the binary part wins)", () => {
    expect(detect(makeZip([{ name: "xl/workbook.xml" }, { name: "xl/workbook.bin" }]), limits)).toBe("xlsb")
    expect(detect(makeZip([{ name: "xl/workbook.bin" }, { name: "xl/workbook.xml" }]), limits)).toBe("xlsb")
  })

  it("truncated input and entries are MalformedFile", () => {
    const zip = makeZip([{ name: "xl/workbook.xml" }])
    for (let cut = 0; cut < zip.length; cut++) {
      const result = readDirectory(zip.subarray(0, cut), limits)
      expect(Array.isArray(result)).toBe(false)
    }
    expect(reason(readDirectory(new Uint8Array(10), limits))).toMatch(/not found/)
  })
})

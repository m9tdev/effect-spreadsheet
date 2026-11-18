/**
 * The input probe for compound files and bare BIFF streams: detection,
 * `Spreadsheet.read` and `Xls.read` apply the same rules, and a read opens
 * the compound file once.
 */
import { Effect, Exit, Stream } from "effect"
import * as Result from "effect/Result"
import { describe, expect, it } from "vitest"
import { probe } from "../src/internal/probe.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { patch32, sectorOffset, writeCfb } from "./helpers/cfb.ts"
import { counter, metered } from "./helpers/meter.ts"
import { failureOf } from "./helpers/run.ts"
import * as X from "./helpers/xls.ts"

const limits = SpreadsheetLimits.defaults
const run = <A, E>(effect: Effect.Effect<A, E>) => Effect.runPromise(Effect.exit(effect))
const formatOf = (bytes: Uint8Array) => Effect.runPromise(Spreadsheet.detectFormat(bytes))
const xlsError = async (bytes: Uint8Array) => failureOf(await run(Xls.read(bytes)))
const readError = async (bytes: Uint8Array) => failureOf(await run(Spreadsheet.read(bytes)))
const tagOf = (bytes: Uint8Array) => {
  const probed = probe(bytes, limits)
  return Result.isSuccess(probed) ? probed.success._tag : probed.failure._tag
}

const book = (): Uint8Array => X.workbookStream({ sheets: [{ name: "S", records: [X.number(0, 0, 0, 7)] }] })

/** A bare stream whose first record is a BOF in all but `patch`. */
const bof = (patch: ReadonlyArray<number>) => {
  const bytes = book()
  bytes.set(patch, 0)
  return bytes
}

describe("probe: bare BIFF streams", () => {
  it("a well-formed BIFF5/8 BOF is a workbook stream, read by both readers", async () => {
    const bytes = book()
    expect(tagOf(bytes)).toBe("Biff")
    expect(await formatOf(bytes)).toBe("xls")
    for (const read of [Xls.read(bytes), Spreadsheet.read(bytes)]) {
      const exit = await run(Effect.flatMap(read, (wb) => Stream.runCollect(Workbook.rows(wb, 0))))
      expect(Exit.isSuccess(exit) && Array.from(exit.value).length).toBe(1)
    }
  })

  it("Xls.read rejects a stream starting 09 08 without a well-formed BOF at once, as detection does", async () => {
    // a BOF of the wrong length, unknown versions (0x0700; 0x0601: BIFF8 is exactly 0x0600), an unknown data type
    const patches = [
      [0x09, 0x08, 0x0C, 0x00],
      [0x09, 0x08, 0x10, 0x00, 0x00, 0x07],
      [0x09, 0x08, 0x10, 0x00, 0x01, 0x06],
      [0x09, 0x08, 0x10, 0x00, 0x00, 0x06, 0x99, 0x00]
    ]
    for (const patch of patches) {
      const bytes = bof(patch)
      expect(await xlsError(bytes), `${patch}`).toMatchObject({
        _tag: "MalformedFile",
        format: "xls",
        reason: "neither a compound file nor a BIFF stream"
      })
      expect(await formatOf(bytes)).not.toBe("xls")
      expect(tagOf(bytes)).toBe("Detected")
    }
  })

  it("a well-formed BIFF2-4 BOF is xls to detection, and UnsupportedFormat naming the version to both readers", async () => {
    const cases: Array<[string, ReadonlyArray<number>]> = [
      ["biff2", [0x09, 0x00, 0x04, 0x00, 0x02, 0x00, 0x10, 0x00]],
      ["biff3", [0x09, 0x02, 0x06, 0x00, 0x00, 0x00, 0x10, 0x00]],
      ["biff4", [0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x10, 0x00]]
    ]
    for (const [version, head] of cases) {
      const bytes = new Uint8Array([...head, 0x0A, 0x00, 0x00, 0x00])
      expect(tagOf(bytes)).toBe("Unreadable")
      expect(await formatOf(bytes)).toBe("xls")
      for (const error of [await xlsError(bytes), await readError(bytes)]) {
        expect(error).toMatchObject({ _tag: "UnsupportedFormat", format: version })
      }
    }
  })

  it("a BIFF2-4 BOF needs no particular version word: the versions files carry still read as before", async () => {
    // version words seen in real files: 0x0007 (BIFF2), 0x0003 (BIFF3), 0x0400 with data type 0x0100 (BIFF4 workbook)
    const cases: Array<[string, ReadonlyArray<number>]> = [
      ["biff2", [0x09, 0x00, 0x04, 0x00, 0x07, 0x00, 0x10, 0x00]],
      ["biff3", [0x09, 0x02, 0x06, 0x00, 0x03, 0x00, 0x10, 0x00]],
      ["biff4", [0x09, 0x04, 0x06, 0x00, 0x00, 0x04, 0x00, 0x01]]
    ]
    for (const [version, head] of cases) {
      const bytes = new Uint8Array([...head, 0x0A, 0x00, 0x00, 0x00])
      expect(tagOf(bytes)).toBe("Unreadable")
      expect(await formatOf(bytes)).toBe("xls")
      for (const error of [await xlsError(bytes), await readError(bytes)]) {
        expect(error).toMatchObject({ _tag: "UnsupportedFormat", format: version })
      }
    }
  })

  it("a stream starting with a BIFF3 or BIFF4 BOF type but no well-formed BOF is MalformedFile at once, to detection and both readers", async () => {
    const cases: Array<[string, string, Uint8Array]> = [
      ["BIFF3", "length 4", new Uint8Array([0x09, 0x02, 0x04, 0x00, 0x00, 0x00, 0x10, 0x00])],
      ["BIFF3", "data type 0x0000", new Uint8Array([0x09, 0x02, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00])],
      ["BIFF3", "two bytes", new Uint8Array([0x09, 0x02])],
      ["BIFF4", "length 8", new Uint8Array([0x09, 0x04, 0x08, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00])],
      ["BIFF4", "data type 0x0033", new Uint8Array([0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x33, 0x00])],
      ["BIFF4", "a BOF cut short", new Uint8Array([0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x10])]
    ]
    for (const [version, label, bytes] of cases) {
      const expected = { _tag: "MalformedFile", format: "xls", reason: `malformed ${version} BOF record` }
      expect(tagOf(bytes), label).toBe("MalformedFile")
      expect(failureOf(await run(Spreadsheet.detectFormat(bytes))), label).toMatchObject(expected)
      expect(await xlsError(bytes), label).toMatchObject(expected)
      expect(await readError(bytes), label).toMatchObject(expected)
    }
  })

  it("bytes starting 09 00 without a well-formed BIFF2 BOF are MalformedFile to Xls.read, and text or unknown data to detection", async () => {
    const utf16le = (text: string) => new Uint8Array([...text].flatMap((c) => [c.charCodeAt(0), 0]))
    const expected = { _tag: "MalformedFile", format: "xls", reason: "malformed BIFF2 BOF record" }
    // UTF-16LE text without a BOM that starts with a tab: its first record looks like a BIFF2 BOF of length 9
    const text = utf16le("\t\tname\tx\n1\t2\t3\n")
    expect(tagOf(text)).toBe("Detected")
    expect(await formatOf(text)).toBe("csv")
    const exit = await run(Effect.flatMap(Spreadsheet.read(text), (wb) => Stream.runCollect(Workbook.rows(wb, 0))))
    expect(Exit.isSuccess(exit) && Array.from(exit.value).length).toBe(2)
    expect(await xlsError(text)).toMatchObject(expected)
    // binary data: a BIFF2 BOF of length 6, one of an unknown data type, two bytes
    const binary = [
      new Uint8Array([0x09, 0x00, 0x06, 0x00, 0x02, 0x00, 0x10, 0x00, 0x00, 0x00]),
      new Uint8Array([0x09, 0x00, 0x04, 0x00, 0x02, 0x00, 0x33, 0x00, 0x00, 0x80, 0x00, 0x90]),
      new Uint8Array([0x09, 0x00])
    ]
    for (const bytes of binary) {
      expect(tagOf(bytes), `${bytes}`).toBe("Detected")
      expect(await formatOf(bytes), `${bytes}`).toBe("unknown")
      expect(await readError(bytes), `${bytes}`).toMatchObject({ _tag: "UnsupportedFormat", format: "unknown" })
      expect(await xlsError(bytes), `${bytes}`).toMatchObject(expected)
    }
  })
})

describe("probe: compound files", () => {
  it("a Workbook storage is no workbook stream: Encrypted with an encrypted package, UnsupportedFormat without", async () => {
    const storage = { name: "Workbook", children: [{ name: "Inner", data: new Uint8Array(100) }] }
    const encrypted = writeCfb([storage, { name: "EncryptionInfo", data: new Uint8Array(200) }, { name: "EncryptedPackage", data: new Uint8Array(5000) }]).bytes
    const other = writeCfb([storage]).bytes
    expect(tagOf(encrypted)).toBe("Encrypted")
    expect(tagOf(other)).toBe("Unreadable")
    for (const bytes of [encrypted, other]) expect(await formatOf(bytes)).toBe("cfb")
    for (const error of [await xlsError(encrypted), await readError(encrypted)]) {
      expect(error).toMatchObject({ _tag: "Encrypted", format: "xlsx" })
    }
    for (const error of [await xlsError(other), await readError(other)]) {
      expect(error).toMatchObject({
        _tag: "UnsupportedFormat",
        format: "cfb",
        reason: "compound file without a Workbook or Book stream"
      })
    }
  })

  it("a workbook stream wins over an encrypted package in the same compound file", async () => {
    const bytes = writeCfb([
      { name: "EncryptionInfo", data: new Uint8Array(200) },
      { name: "EncryptedPackage", data: new Uint8Array(5000) },
      { name: "Workbook", data: book() }
    ]).bytes
    expect(tagOf(bytes)).toBe("Biff")
    expect(await formatOf(bytes)).toBe("xls")
    for (const read of [Xls.read(bytes), Spreadsheet.read(bytes)]) {
      const exit = await run(read)
      expect(Exit.isSuccess(exit) && exit.value.sheets.map((s) => s.name)).toEqual(["S"])
    }
  })

  it("a Workbook stream wins over a Book stream", async () => {
    const biff5 = X.workbookStream({ biff: 5, sheets: [{ name: "FromBook", records: [] }] })
    const biff8 = X.workbookStream({ sheets: [{ name: "FromWorkbook", records: [] }] })
    const bytes = writeCfb([{ name: "Book", data: biff5 }, { name: "Workbook", data: biff8 }]).bytes
    for (const read of [Xls.read(bytes), Spreadsheet.read(bytes)]) {
      const exit = await run(read)
      expect(Exit.isSuccess(exit) && exit.value.sheets.map((s) => s.name)).toEqual(["FromWorkbook"])
    }
  })

  it("detection reads the directory only: a broken workbook stream is still xls, and fails when read", async () => {
    const stream = new Uint8Array(8192) // past the mini-stream cutoff: in regular sectors
    stream.set(book())
    const layout = writeCfb([{ name: "Workbook", data: stream }])
    const { start } = layout.streams.get("Workbook")!
    // the stream's first FAT entry points past the file
    patch32(layout.bytes, sectorOffset(layout.fatSectors[0]!) + start * 4, 0x00FFFFF0)
    expect(tagOf(layout.bytes)).toBe("Biff")
    expect(await formatOf(layout.bytes)).toBe("xls")
    for (const error of [await xlsError(layout.bytes), await readError(layout.bytes)]) {
      expect(error).toMatchObject({ _tag: "MalformedFile", format: "cfb" })
    }
  })

  it("Spreadsheet.read opens the compound file once, metered, as Xls.read does", async () => {
    const stream = new Uint8Array(20_000)
    stream.set(book())
    const bytes = writeCfb([{ name: "Workbook", data: stream }]).bytes
    const sectors = async (read: Effect.Effect<Workbook.Workbook, unknown>) => {
      const meter = counter("sectors")
      await Effect.runPromise(Effect.ignore(metered(meter, read)))
      return meter.total()
    }
    const xls = await sectors(Xls.read(bytes))
    expect(xls).toBeGreaterThan(40) // the directory and FAT chains, and the stream's sectors
    expect(await sectors(Spreadsheet.read(bytes))).toBe(xls)
  })
})

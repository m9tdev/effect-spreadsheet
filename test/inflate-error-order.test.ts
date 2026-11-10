import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { HOLD_BYTES, inflateRawWeb } from "../src/internal/zip/inflate.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsb from "../src/xlsb/index.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import * as B from "./helpers/xlsb.ts"
import { buildZip, concat, deflateRaw, testCrc32 } from "./helpers/zip.ts"

const enc = (s: string) => new TextEncoder().encode(s)
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`
const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

/** Deflate data that ends with 4 bytes after the end of the stream, or with its last 20 bytes cut. */
const broken = async (data: Uint8Array, how: "trailing" | "truncated") => {
  const body = await deflateRaw(data)
  return how === "trailing" ? concat([body, new Uint8Array([1, 2, 3, 4])]) : body.subarray(0, body.length - 20)
}

/** A one-sheet package whose sheet part is `data` deflated, then broken. */
const pkg = async (format: "xlsb" | "xlsx", data: Uint8Array, how: "trailing" | "truncated") => {
  const sheet = {
    name: `xl/worksheets/sheet1.${format === "xlsb" ? "bin" : "xml"}`,
    data: await broken(data, how),
    method: "store" as const,
    methodCode: 8,
    crc32: testCrc32(data),
    uncompressedSize: data.length
  }
  const rels = (target: string) =>
    enc(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/${target}"/></Relationships>`)
  return buildZip(format === "xlsb"
    ? [
      { name: "_rels/.rels", data: rels(`officeDocument" Target="xl/workbook.bin`) },
      { name: "xl/workbook.bin", data: B.concat([B.beginBook(), B.wbProp(false), B.record(0x8F), B.bundleSh("S", "rId1"), B.record(0x90), B.endBook()]) },
      sheet,
      { name: "xl/_rels/workbook.bin.rels", data: rels(`worksheet" Target="worksheets/sheet1.bin`) }
    ]
    : [
      { name: "_rels/.rels", data: rels(`officeDocument" Target="xl/workbook.xml`) },
      { name: "xl/workbook.xml", data: enc(`<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`) },
      sheet,
      { name: "xl/_rels/workbook.xml.rels", data: rels(`worksheet" Target="worksheets/sheet1.xml`) }
    ])
}

/** 2000 rows of one number each (about 66 KB as XLSB, 93 KB as XLSX: several inflate chunks). */
const ROWS = 2000
const xlsbSheet = () => {
  const records: Array<Uint8Array> = []
  for (let r = 0; r < ROWS; r++) records.push(B.rowHdr(r), B.rk(0, B.rkInt(r)))
  return B.sheetBin(records)
}
const xlsxSheet = () => {
  let rows = ""
  for (let r = 0; r < ROWS; r++) rows += `<row r="${r + 1}"><c r="A${r + 1}"><v>${r}</v></c></row>`
  return enc(`<?xml version="1.0" encoding="UTF-8"?><worksheet ${NS}><sheetData>${rows}</sheetData></worksheet>`)
}

/** The tag (and limit) a full read of the sheet fails with; `slow` waits a timer every 100 rows. */
const outcome = async (format: "xlsb" | "xlsx", bytes: Uint8Array, slow: boolean) => {
  const read = format === "xlsb" ? Xlsb.read : Xlsx.read
  const exit = await Effect.runPromiseExit(Effect.flatMap(read(bytes, { limits: { maxRows: ROWS - 100 } }), (workbook) =>
    Stream.runDrain(Workbook.rows(workbook, 0).pipe(
      Stream.tap((row) => (slow && row.index % 100 === 0 ? Effect.promise(tick) : Effect.void))
    ))))
  if (Exit.isSuccess(exit)) return "success"
  const error = exit.cause.reasons.find((reason) => reason._tag === "Fail")?.error as { _tag: string; limit?: string } | undefined
  return `${error?._tag} ${error?.limit ?? ""}`.trim()
}

describe("an inflate error after a parse error", () => {
  for (const how of ["trailing", "truncated"] as const) {
    it(`inflateRawWeb returns every byte decompressed before the error, however slowly it is read (${how})`, async () => {
      const data = xlsbSheet()
      const compressed = await broken(data, how)
      // progress: one read per run
      for (let run = 0; run < 5; run++) {
        const inflater = inflateRawWeb(compressed)
        let size = 0
        let failed = false
        // progress: each next() returns a chunk, the end, or rejects
        while (true) {
          const chunk = await inflater.next().catch(() => {
            failed = true
            return undefined
          })
          if (chunk === undefined) break
          size += chunk.length
          await tick()
        }
        expect(failed).toBe(true)
        // trailing junk: the whole stream decoded; truncated: all but what the cut bytes held
        if (how === "trailing") expect(size).toBe(data.length)
        else expect(size).toBeGreaterThan(data.length - 16 * 1024)
      }
    })
  }

  for (const format of ["xlsb", "xlsx"] as const) {
    for (const how of ["trailing", "truncated"] as const) {
      it(`${format}: the earlier maxRows error wins over broken deflate data (${how}), on every read and at any reading speed`, async () => {
        const bytes = await pkg(format, format === "xlsb" ? xlsbSheet() : xlsxSheet(), how)
        const seen = new Map<string, number>()
        // progress: run up to 40
        for (let run = 0; run < 40; run++) {
          const key = await outcome(format, bytes, run % 4 === 0)
          seen.set(key, (seen.get(key) ?? 0) + 1)
        }
        expect(Object.fromEntries(seen)).toEqual({ "LimitExceeded maxRows": 40 })
      }, 60_000)
    }
  }

  it("inflateRawWeb holds at most about HOLD_BYTES of output ahead of a slow reader, even for a deflate bomb", async () => {
    // 64 MiB of zeros deflate to about one 64 KiB slice: a slice that inflates over a thousandfold
    const total = 64 * 1024 * 1024
    const compressed = await deflateRaw(new Uint8Array(total))
    expect(compressed.length).toBeLessThanOrEqual(64 * 1024)
    // Counts the bytes taken off the decompressor's readable side, one chunk per read.
    let taken = 0
    const Real = globalThis.DecompressionStream
    class Counting {
      readonly writable: WritableStream<BufferSource>
      readonly readable: ReadableStream<Uint8Array>
      constructor(format: CompressionFormat) {
        const inner = new Real(format)
        const source = inner.readable.getReader()
        this.writable = inner.writable
        this.readable = new ReadableStream<Uint8Array>({
          pull: async (controller) => {
            const result = await source.read()
            if (result.done) controller.close()
            else {
              taken += result.value.length
              controller.enqueue(result.value)
            }
          },
          cancel: (reason) => source.cancel(reason)
        }, { highWaterMark: 0 })
      }
    }
    globalThis.DecompressionStream = Counting as unknown as typeof DecompressionStream
    try {
      const inflater = inflateRawWeb(compressed)
      let handed = 0
      let most = 0
      // progress: each next() returns a chunk or the end; the output is finite
      for (let reads = 0;; reads++) {
        const chunk = await inflater.next()
        if (chunk === undefined) break
        handed += chunk.length
        // give the decompressor time to run ahead, now and then
        if (reads % 512 === 0) for (let i = 0; i < 20; i++) await tick()
        most = Math.max(most, taken - handed)
      }
      expect(handed).toBe(total)
      // the pump may take one more chunk (at most 64 KiB) before it sees the cap
      expect(most).toBeLessThanOrEqual(HOLD_BYTES + 64 * 1024)
      expect(most).toBeGreaterThan(HOLD_BYTES / 2)
    } finally {
      globalThis.DecompressionStream = Real
    }
  }, 60_000)
})

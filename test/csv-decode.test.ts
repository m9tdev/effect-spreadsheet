import { describe, expect, it } from "@effect/vitest"
import { Stream } from "effect"
import { ByteDecoder, DETECT_BYTES } from "../src/internal/csv/decode.ts"
import { InvalidOption } from "../src/SpreadsheetError.ts"
import { chunksOf, csvUntilError, enc } from "./helpers/run.ts"

const utf8Bom = new Uint8Array([0xEF, 0xBB, 0xBF])

describe("csv: byte decoder", () => {
  it("a BOM split across pushes (1 + 2 bytes) is still found and dropped, auto or requested", () => {
    // é in windows-1252 would be "Ã©"; the UTF-8 BOM has to win
    const body = enc("né")
    const auto = new ByteDecoder("auto")
    expect(auto.push(utf8Bom.subarray(0, 1))).toBe("")
    expect(auto.push(utf8Bom.subarray(1))).toBe("")
    expect(auto.push(body)).toBe("")
    expect(auto.end()).toBe("né")

    const requested = new ByteDecoder("windows-1252")
    expect(requested.push(utf8Bom.subarray(0, 1))).toBe("")
    // 3 bytes held: the encoding is decided here
    expect(requested.push(utf8Bom.subarray(1))).toBe("")
    expect(requested.push(body)).toBe("né")
    expect(requested.end()).toBe("")
  })

  it("auto detection waits for exactly DETECT_BYTES bytes, then decodes all of them", () => {
    const bytes = new Uint8Array(DETECT_BYTES).fill(0x61)
    const decoder = new ByteDecoder("auto")
    expect(decoder.push(bytes.subarray(0, DETECT_BYTES - 1))).toBe("")
    const text = decoder.push(bytes.subarray(DETECT_BYTES - 1))
    expect(typeof text === "string" && text.length).toBe(DETECT_BYTES)
    expect(decoder.push(enc("b"))).toBe("b")
    expect(decoder.end()).toBe("")

    const oneChunk = new ByteDecoder("auto")
    const all = oneChunk.push(bytes)
    expect(typeof all === "string" && all.length).toBe(DETECT_BYTES)
  })

  it("input shorter than the detection window is decoded at the end", () => {
    const decoder = new ByteDecoder("auto")
    expect(decoder.push(enc("a,"))).toBe("")
    expect(decoder.push(enc("b"))).toBe("")
    expect(decoder.end()).toBe("a,b")
    expect(new ByteDecoder("auto").end()).toBe("")
  })

  it("an unknown requested encoding is an error, and every later push or end returns that same error", () => {
    const decoder = new ByteDecoder("klingon")
    expect(decoder.push(enc("a"))).toBe("")
    const error = decoder.push(enc("bc"))
    expect(error).toBeInstanceOf(InvalidOption)
    expect(decoder.push(enc("d"))).toBe(error)
    expect(decoder.end()).toBe(error)

    const atEnd = new ByteDecoder("klingon")
    expect(atEnd.push(enc("a"))).toBe("")
    const endError = atEnd.end()
    expect(endError).toBeInstanceOf(InvalidOption)
    expect(atEnd.end()).toBe(endError)
  })

  it("an unknown encoding through rows on a multi-chunk stream fails with InvalidOption and no rows", async () => {
    const input = Stream.fromIterable(chunksOf(enc("a\nb\nc\n"), 1))
    const { error, rows } = await csvUntilError(input, { encoding: "klingon" })
    expect(rows).toEqual([])
    expect(error?._tag === "InvalidOption" && error.option).toBe("encoding")
  })
})

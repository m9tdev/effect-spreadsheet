import { describe, expect, it } from "@effect/vitest"
import { Stream } from "effect"
import * as Result from "effect/Result"
import { ByteDecoder, DETECT_BYTES } from "../src/internal/csv/decode.ts"
import { type Decoding, parseEncoding } from "../src/internal/csv/options.ts"
import { InvalidOption } from "../src/SpreadsheetError.ts"
import { chunksOf, csvUntilError, enc } from "./helpers/run.ts"

const utf8Bom = new Uint8Array([0xEF, 0xBB, 0xBF])

/** The text of a decoder step. */
const textOf = (text: string): string => text

/** The decoding for a label that must be known. */
const decoding = (label: string): Decoding => {
  const parsed = parseEncoding(label)
  if (Result.isFailure(parsed)) throw new Error(`unexpected ${parsed.failure._tag}`)
  return parsed.success
}

describe("csv: byte decoder", () => {
  it("a BOM split across pushes (1 + 2 bytes) is still found and dropped, auto or requested", () => {
    // é in windows-1252 would be "Ã©"; the UTF-8 BOM has to win
    const body = enc("né")
    const auto = new ByteDecoder("auto")
    expect(textOf(auto.push(utf8Bom.subarray(0, 1)))).toBe("")
    expect(textOf(auto.push(utf8Bom.subarray(1)))).toBe("")
    expect(textOf(auto.push(body))).toBe("")
    expect(textOf(auto.end())).toBe("né")

    const requested = new ByteDecoder(decoding("windows-1252"))
    expect(textOf(requested.push(utf8Bom.subarray(0, 1)))).toBe("")
    // 3 bytes held: the encoding is decided here
    expect(textOf(requested.push(utf8Bom.subarray(1)))).toBe("")
    expect(textOf(requested.push(body))).toBe("né")
    expect(textOf(requested.end())).toBe("")
  })

  it("auto detection waits for exactly DETECT_BYTES bytes, then decodes all of them", () => {
    const bytes = new Uint8Array(DETECT_BYTES).fill(0x61)
    const decoder = new ByteDecoder("auto")
    expect(textOf(decoder.push(bytes.subarray(0, DETECT_BYTES - 1)))).toBe("")
    const text = textOf(decoder.push(bytes.subarray(DETECT_BYTES - 1)))
    expect(text.length).toBe(DETECT_BYTES)
    expect(textOf(decoder.push(enc("b")))).toBe("b")
    expect(textOf(decoder.end())).toBe("")

    const oneChunk = new ByteDecoder("auto")
    const all = textOf(oneChunk.push(bytes))
    expect(all.length).toBe(DETECT_BYTES)
  })

  it("input shorter than the detection window is decoded at the end", () => {
    const decoder = new ByteDecoder("auto")
    expect(textOf(decoder.push(enc("a,")))).toBe("")
    expect(textOf(decoder.push(enc("b")))).toBe("")
    expect(textOf(decoder.end())).toBe("a,b")
    expect(textOf(new ByteDecoder("auto").end())).toBe("")
  })

  it("an unknown encoding is an InvalidOption when parsed, with the label in the reason; known labels make fresh decoders", () => {
    const unknown = parseEncoding("klingon")
    if (Result.isSuccess(unknown)) throw new Error("expected an error")
    expect(unknown.failure).toBeInstanceOf(InvalidOption)
    expect(unknown.failure.option).toBe("encoding")
    expect(unknown.failure.reason).toBe('text encoding "klingon" is not supported')
    expect(parseEncoding("auto")).toEqual(Result.succeed("auto"))
    const latin1 = decoding("latin1")
    if (latin1 === "auto") throw new Error("expected a decoder maker")
    const a = latin1()
    const b = latin1()
    expect(a).not.toBe(b)
    expect(a.encoding).toBe("windows-1252")
    expect(a.ignoreBOM).toBe(false)
  })

  it("an unknown encoding through rows on a multi-chunk stream fails with InvalidOption and no rows", async () => {
    const input = Stream.fromIterable(chunksOf(enc("a\nb\nc\n"), 1))
    const { error, rows } = await csvUntilError(input, { encoding: "klingon" })
    expect(rows).toEqual([])
    expect(error?._tag === "InvalidOption" && error.option).toBe("encoding")
  })
})

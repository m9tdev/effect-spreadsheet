import { describe, expect, it } from "@effect/vitest"
import { Stream } from "effect"
import * as Result from "effect/Result"
import { ByteDecoder, DETECT_BYTES } from "../src/internal/csv/decode.ts"
import { InvalidOption } from "../src/SpreadsheetError.ts"
import { chunksOf, csvUntilError, enc } from "./helpers/run.ts"

const utf8Bom = new Uint8Array([0xEF, 0xBB, 0xBF])

/** The text of a decoder step that must succeed. */
const textOf = (r: Result.Result<string, InvalidOption>): string => {
  if (Result.isFailure(r)) throw new Error(`unexpected ${r.failure._tag}`)
  return r.success
}

/** The error of a decoder step that must fail. */
const errorOf = (r: Result.Result<string, InvalidOption>): InvalidOption => {
  if (Result.isSuccess(r)) throw new Error(`expected an error, got ${JSON.stringify(r.success)}`)
  return r.failure
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

    const requested = new ByteDecoder("windows-1252")
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

  it("an unknown requested encoding is an error, and every later push or end returns that same error", () => {
    const decoder = new ByteDecoder("klingon")
    expect(textOf(decoder.push(enc("a")))).toBe("")
    const error = errorOf(decoder.push(enc("bc")))
    expect(error).toBeInstanceOf(InvalidOption)
    expect(errorOf(decoder.push(enc("d")))).toBe(error)
    expect(errorOf(decoder.end())).toBe(error)

    const atEnd = new ByteDecoder("klingon")
    expect(textOf(atEnd.push(enc("a")))).toBe("")
    const endError = errorOf(atEnd.end())
    expect(endError).toBeInstanceOf(InvalidOption)
    expect(errorOf(atEnd.end())).toBe(endError)
  })

  it("an unknown encoding through rows on a multi-chunk stream fails with InvalidOption and no rows", async () => {
    const input = Stream.fromIterable(chunksOf(enc("a\nb\nc\n"), 1))
    const { error, rows } = await csvUntilError(input, { encoding: "klingon" })
    expect(rows).toEqual([])
    expect(error?._tag === "InvalidOption" && error.option).toBe("encoding")
  })
})

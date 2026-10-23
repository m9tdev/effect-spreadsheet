import { describe, it } from "@effect/vitest"
import { Stream } from "effect"
import fc from "fast-check"
import { expect } from "vitest"
import { csvAoa, enc, splitAt } from "./helpers/run.ts"

/** Independent, naive reference for the lenient parser (whole string at once). */
const reference = (text: string, delimiter: string): Array<Array<string | undefined>> => {
  const rows: Array<Array<string | undefined>> = []
  let row: Array<string | undefined> = []
  let field = ""
  let started = false
  let i = 0
  const pushField = () => {
    row.push(field === "" ? undefined : field)
    field = ""
  }
  const pushRow = () => {
    // trailing empty fields produce no cells
    while (row.length > 0 && row[row.length - 1] === undefined) row.pop()
    rows.push(row)
    row = []
    started = false
  }
  while (i < text.length) {
    started = true
    if (text[i] === '"') {
      i++
      while (i < text.length) {
        if (text[i] === '"') {
          if (text[i + 1] === '"') {
            field += '"'
            i += 2
          } else {
            i++
            break
          }
        } else field += text[i++]
      }
    }
    while (i < text.length && text[i] !== delimiter && text[i] !== "\n" && text[i] !== "\r") field += text[i++]
    if (i >= text.length) break
    const c = text[i++]
    pushField()
    if (c !== delimiter) {
      if (c === "\r" && text[i] === "\n") i++
      pushRow()
    }
  }
  if (started) {
    pushField()
    pushRow()
  }
  return rows
}

const csvText = fc
  .array(fc.constantFrom("a", "b", "1", " ", "é", "€", "😀", ",", ",", ";", '"', '"', "\r", "\n", "\n", "\r\n", '""'), {
    maxLength: 200
  })
  .map((parts) => parts.join(""))

const dense = (aoa: Array<Array<unknown>>) =>
  aoa.map((row) => {
    const out = [...row]
    while (out.length > 0 && out[out.length - 1] === undefined) out.pop()
    return out
  })

const cutsFor = (length: number) => fc.array(fc.integer({ min: 0, max: Math.max(0, length) }), { maxLength: 12 })

describe("csv properties", () => {
  it("matches a naive reference parser", async () => {
    await fc.assert(
      fc.asyncProperty(csvText, async (text) => {
        const ours = await csvAoa(text, { delimiter: "," })
        expect(dense(ours)).toEqual(reference(text, ","))
      }),
      { numRuns: 300 }
    )
  })

  it("any chunking of the bytes gives identical rows (multibyte chars, quotes, CRLF split anywhere)", async () => {
    await fc.assert(
      fc.asyncProperty(
        csvText.chain((text) => fc.tuple(fc.constant(text), cutsFor(enc(text).length))),
        fc.boolean(),
        async ([text, cuts], inferTypes) => {
          const bytes = enc(text)
          const whole = await csvAoa(bytes, { inferTypes })
          const chunked = await csvAoa(Stream.fromIterable(splitAt(bytes, cuts)), { inferTypes })
          expect(chunked).toEqual(whole)
        }
      ),
      { numRuns: 300 }
    )
  })

  it("chunking invariance also holds with UTF-16 and a BOM", async () => {
    await fc.assert(
      fc.asyncProperty(csvText.chain((text) => fc.tuple(fc.constant(text), cutsFor(text.length * 2 + 2))), async ([text, cuts]) => {
        const bytes = new Uint8Array(text.length * 2 + 2)
        bytes[0] = 0xFF
        bytes[1] = 0xFE
        for (let i = 0; i < text.length; i++) {
          bytes[2 + 2 * i] = text.charCodeAt(i) & 0xFF
          bytes[3 + 2 * i] = text.charCodeAt(i) >> 8
        }
        const whole = await csvAoa(bytes)
        const chunked = await csvAoa(Stream.fromIterable(splitAt(bytes, cuts)))
        expect(chunked).toEqual(whole)
        expect(whole).toEqual(await csvAoa(text))
      }),
      { numRuns: 150 }
    )
  })

  it("chunking invariance across the 64 KiB detection window and 16 Ki sniff window", async () => {
    const big = ('x;"y\r\n";é€\r\n').repeat(9000) // ~117 KB, ; delimiter by sniffing
    const bytes = enc(big)
    const whole = await csvAoa(bytes)
    expect(whole.length).toBe(9000)
    expect(whole[0]).toEqual(["x", "y\r\n", "é€"])
    await fc.assert(
      fc.asyncProperty(cutsFor(bytes.length), async (cuts) => {
        expect(await csvAoa(Stream.fromIterable(splitAt(bytes, cuts)))).toEqual(whole)
      }),
      { numRuns: 20 }
    )
    // one byte at a time
    const oneByOne = Stream.fromIterable(Array.from({ length: 3000 }, (_, i) => bytes.subarray(i, i + 1))).pipe(
      Stream.concat(Stream.make(bytes.subarray(3000)))
    )
    expect(await csvAoa(oneByOne)).toEqual(whole)
  })
})

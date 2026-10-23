import { describe, it } from "@effect/vitest"
import { Effect, Stream } from "effect"
import fc from "fast-check"
import { expect } from "vitest"
import * as Csv from "../src/csv/index.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as Workbook from "../src/Workbook.ts"
import { classify, csvAoa, enc, splitAt } from "./helpers/run.ts"
import { makeZip } from "./helpers/zip.ts"

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

const BUDGET_MS = 2000

/** Runs `effect` and checks it ends with a value or a typed error, within the budget. */
const terminates = async (effect: Effect.Effect<unknown, unknown>) => {
  const start = performance.now()
  const exit = await Effect.runPromise(Effect.exit(effect))
  const elapsed = performance.now() - start
  const result = classify(exit)
  expect(result.detail ?? "", "no defect").toBe("")
  expect(result.kind).not.toBe("defect")
  expect(elapsed).toBeLessThan(BUDGET_MS)
}

const readAll = (bytes: Uint8Array, limits?: import("../src/SpreadsheetLimits.ts").LimitOverrides) =>
  Effect.flatMap(Spreadsheet.read(bytes, { limits }), (wb) =>
    Effect.forEach(wb.sheets, (sheet) => Stream.runDrain(Workbook.rows(wb, sheet.index))))

const SAMPLES = [
  enc('a,b,"c\r\nd"\r\n1,2,3\n'),
  enc("sep=;\nx;y\n"),
  makeZip([{ name: "xl/workbook.xml" }, { name: "xl/worksheets/sheet1.xml", data: enc("<worksheet/>") }]),
  new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, ...new Uint8Array(100)]),
  enc('<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"/>'),
  new Uint8Array([0xFF, 0xFE, 0x61, 0x00, 0x2C, 0x00, 0x22, 0x00])
]

describe("fuzz: never a defect, never a hang", () => {
  it("random bytes through Spreadsheet.read and Csv.rows (default and tiny limits)", async () => {
    await fc.assert(
      fc.asyncProperty(fc.uint8Array({ maxLength: 4096 }), fc.boolean(), async (bytes, tiny) => {
        const limits = tiny ? { maxColumns: 3, maxRows: 5, maxStringLength: 8, maxCellsPerSheet: 7, maxZipEntries: 2 } : undefined
        await terminates(readAll(bytes, limits))
        await terminates(Stream.runDrain(Csv.rows(bytes, { limits, inferTypes: true, strict: tiny })))
      }),
      { numRuns: 300 }
    )
  })

  it("truncations and bit flips of samples", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...SAMPLES).chain((sample) =>
          fc.tuple(
            fc.constant(sample),
            fc.integer({ min: 0, max: sample.length }),
            fc.array(fc.tuple(fc.nat(), fc.integer({ min: 0, max: 7 })), { maxLength: 8 })
          )
        ),
        async ([sample, length, flips]) => {
          const bytes = sample.slice(0, length)
          for (const [at, bit] of flips) if (bytes.length > 0) bytes[at % bytes.length]! ^= 1 << bit
          await terminates(readAll(bytes))
          await terminates(Spreadsheet.detectFormat(bytes))
          await terminates(Stream.runDrain(Csv.rows(bytes, { strict: true })))
        }
      ),
      { numRuns: 400 }
    )
  })

  it("random chunked streams of random bytes", async () => {
    await fc.assert(
      fc.asyncProperty(fc.array(fc.uint8Array({ maxLength: 300 }), { maxLength: 20 }), async (chunks) => {
        await terminates(Stream.runDrain(Csv.rows(Stream.fromIterable(chunks), { inferTypes: true })))
      }),
      { numRuns: 200 }
    )
  })

  it("pathological inputs finish fast", async () => {
    const cases = [
      '"'.repeat(200_000),
      ",".repeat(200_000),
      "\n".repeat(200_000),
      "\r".repeat(200_000),
      '""'.repeat(100_000),
      '"a"b'.repeat(50_000),
      "\r\n".repeat(100_000)
    ]
    for (const text of cases) {
      await terminates(Stream.runDrain(Csv.rows(enc(text))))
      await terminates(Stream.runDrain(Csv.rows(enc(text), { strict: true })))
    }
  })
})

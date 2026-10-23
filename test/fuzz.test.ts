import { readFileSync } from "node:fs"
/**
 * Fuzz: random and mutated input must end in a value or a typed
 * SpreadsheetError, within a time budget. Each case runs in a worker thread
 * that is terminated when it overruns, so a synchronous infinite loop fails
 * the test instead of hanging the run.
 */
import fc from "fast-check"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { FuzzHost } from "./fuzz/host.ts"
import type { FuzzTask } from "./fuzz/tasks.ts"
import { enc } from "./helpers/run.ts"
import { makeZip } from "./helpers/zip.ts"

// Per case, in a worker that is terminated when it overruns. The point is
// hang detection: cases take milliseconds (the 200k-character pathological
// inputs measured 0-13 ms each, warm, on an idle machine), while a hang
// never ends and quadratic work on those inputs takes far longer. 5 s
// leaves room for a loaded machine.
const BUDGET_MS = 5000
const host = new FuzzHost(BUDGET_MS)

beforeAll(() => host.start(), 60_000)
afterAll(() => host.stop())

const terminates = async (task: FuzzTask) => {
  const result = await host.run(task)
  expect(result.kind === "defect" || result.kind === "hang" ? `${result.kind}: ${result.detail}` : "ok").toBe("ok")
}

const SAMPLES = [
  enc('a,b,"c\r\nd"\r\n1,2,3\n'),
  enc("sep=;\nx;y\n"),
  makeZip([{ name: "xl/workbook.xml" }, { name: "xl/worksheets/sheet1.xml", data: enc("<worksheet/>") }]),
  new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, ...new Uint8Array(100)]),
  enc('<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"/>'),
  new Uint8Array([0xFF, 0xFE, 0x61, 0x00, 0x2C, 0x00, 0x22, 0x00])
]

describe("fuzz: never a defect, never a hang", () => {
  it("the harness itself reports a hang as a failure", async () => {
    // A task that can't finish in 1 ms of budget is reported, not waited for.
    const result = await host.run({ op: "csv", chunks: [enc("a,b\n".repeat(200_000))] }, 1)
    expect(result.kind).toBe("hang")
    // and the next task gets a fresh worker
    expect((await host.run({ op: "detect", bytes: enc("a") }, 30_000)).kind).toBe("value")
  })

  it("random bytes through Spreadsheet.read and Csv.rows (default and tiny limits)", async () => {
    await fc.assert(
      fc.asyncProperty(fc.uint8Array({ maxLength: 4096 }), fc.boolean(), async (bytes, tiny) => {
        const limits = tiny ? { maxColumns: 3, maxRows: 5, maxStringLength: 8, maxCellsPerSheet: 7, maxZipEntries: 2 } : undefined
        await terminates({ op: "read", bytes, limits })
        await terminates({ op: "csv", chunks: [bytes], options: { limits, inferTypes: true, strict: tiny } })
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
          await terminates({ op: "read", bytes })
          await terminates({ op: "detect", bytes })
          await terminates({ op: "csv", chunks: [bytes], options: { strict: true } })
        }
      ),
      { numRuns: 400 }
    )
  })

  it("random chunked streams of random bytes", async () => {
    await fc.assert(
      fc.asyncProperty(fc.array(fc.uint8Array({ maxLength: 300 }), { maxLength: 20 }), async (chunks) => {
        await terminates({ op: "csv", chunks, options: { inferTypes: true } })
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
      await terminates({ op: "csv", chunks: [enc(text)] })
      await terminates({ op: "csv", chunks: [enc(text)], options: { strict: true } })
    }
  })

  it("zip + xml: truncations and bit flips of the corpus workbooks", async () => {
    const workbooks = ["test/corpus/generated/features.xlsx", "test/corpus/generated/strings.xlsx", "test/corpus/hidden-rows/hidden-row.xlsx", "test/corpus/generated/features.ods"]
      .map((p) => new Uint8Array(readFileSync(p)))
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...workbooks).chain((sample) =>
          fc.tuple(
            fc.constant(sample),
            fc.integer({ min: 0, max: sample.length }),
            fc.array(fc.tuple(fc.nat(), fc.integer({ min: 0, max: 7 })), { maxLength: 8 })
          )
        ),
        async ([sample, length, flips]) => {
          const bytes = sample.slice(0, length)
          for (const [at, bit] of flips) if (bytes.length > 0) bytes[at % bytes.length]! ^= 1 << bit
          await terminates({ op: "zip", bytes })
          await terminates({ op: "read", bytes })
        }
      ),
      { numRuns: 300 }
    )
  })

  it("xml: random and mutated documents in random chunks", async () => {
    const sheet = new TextEncoder().encode(
      `<?xml version="1.0"?><worksheet xmlns="x"><sheetData><row r="1" hidden="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>a &amp; b</t></is></c></row></sheetData><!-- c --><mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>`
    )
    await fc.assert(
      fc.asyncProperty(
        fc.oneof(
          fc.uint8Array({ maxLength: 2048 }),
          fc.tuple(fc.integer({ min: 0, max: sheet.length }), fc.array(fc.tuple(fc.nat(), fc.integer({ min: 0, max: 7 })), { maxLength: 6 }))
            .map(([length, flips]) => {
              const bytes = sheet.slice(0, length)
              for (const [at, bit] of flips) if (bytes.length > 0) bytes[at % bytes.length]! ^= 1 << bit
              return bytes
            })
        ),
        fc.array(fc.nat(), { maxLength: 8 }),
        async (bytes, cuts) => {
          const points = [...new Set(cuts.map((c) => c % (bytes.length + 1)))].sort((a, b) => a - b)
          const chunks: Array<Uint8Array> = []
          let last = 0
          for (const p of points) {
            chunks.push(bytes.slice(last, p))
            last = p
          }
          chunks.push(bytes.slice(last))
          await terminates({ op: "xml", chunks, limits: { maxXmlDepth: 8, maxStringLength: 64 } })
        }
      ),
      { numRuns: 500 }
    )
  })
})

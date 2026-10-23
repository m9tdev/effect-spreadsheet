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
// hang detection: cases take milliseconds, the 200k-character pathological
// inputs well under 100 ms, while a hang never ends and quadratic work on
// those inputs takes far longer. 5 s leaves room for a loaded machine.
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
})

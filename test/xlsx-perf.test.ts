/**
 * PLAN.md performance check: a 200 000-row XLSX streams with bounded work
 * between event-loop turns, in time linear in its size. Both are checked
 * without wall-clock limits: the work per turn as a count (rows between
 * turns), and linearity as a ratio to reference work in the same process
 * (tokenizing the same sheet part with a handler that does nothing), which
 * machine load slows alike. scripts/bench/xlsx-rows.ts measures the actual
 * pauses.
 */
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { parse } from "../src/internal/xml/parse.ts"
import { open } from "../src/internal/zip/archive.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { bigXlsx, MIN_ROW_BYTES } from "./helpers/big-xlsx.ts"
import { countYields } from "./helpers/yields.ts"

/** The most bytes of sheet XML one pull handles: one inflated or stored piece. */
const PIECE_BYTES = 64 * 1024
const ROWS = 200_000
/**
 * Reading rows may take at most this many times as long as just
 * tokenizing the part. Measured: 1.5-1.8x idle, 0.9-1.7x with 4 full
 * suites and 8 CPU burners running (load average ~60-75). A reader that
 * rescans all earlier rows measured 22x and fails; one rescanning on every
 * 4th row (5.5x) or less often is near or below what timing can see.
 */
const MAX_RELATIVE = 6

const timed = async (run: () => Promise<unknown>): Promise<number> => {
  const start = performance.now()
  await run()
  return performance.now() - start
}

describe("xlsx: 200 000 rows", () => {
  it("stream with at most one piece of rows between event-loop turns, in linear time", async () => {
    const bytes = await bigXlsx(ROWS)
    const wb = await Effect.runPromise(Xlsx.read(bytes))
    let rows = 0
    let hidden = 0
    let cells = 0
    const run = await countYields(
      () =>
        Effect.runPromise(Stream.runForEach(Workbook.rows(wb, 0), (row) =>
          Effect.sync(() => {
            rows++
            cells += row.cells.length
            if (row.hidden) hidden++
          }))),
      () => rows
    )
    expect(rows).toBe(ROWS)
    expect(hidden).toBe(ROWS / 10)
    expect(cells).toBe(4 * ROWS)
    // the work between two turns is at most one piece of XML worth of rows
    const perPiece = Math.ceil(PIECE_BYTES / MIN_ROW_BYTES) + 1
    expect(run.maxPerTurn).toBeLessThanOrEqual(perPiece)
    expect(run.yields).toBeGreaterThanOrEqual(ROWS / perPiece)

    // linear: compared with tokenizing the same part, best of 3, interleaved
    // in alternating order
    const archive = open(bytes, SpreadsheetLimits.defaults)
    if ("_tag" in archive) throw new Error(archive._tag)
    const entry = archive.find("xl/worksheets/sheet1.xml")!
    const ignore = { open: () => undefined, close: () => undefined, text: () => undefined }
    const reference = () => Effect.runPromise(parse(archive.read(entry), ignore, SpreadsheetLimits.defaults))
    const reading = () => Effect.runPromise(Stream.runDrain(Workbook.rows(wb, 0)))
    let readMs = Infinity
    let referenceMs = Infinity
    // progress: r increases up to 3
    for (let r = 0; r < 3; r++) {
      if (r % 2 === 0) {
        referenceMs = Math.min(referenceMs, await timed(reference))
        readMs = Math.min(readMs, await timed(reading))
      } else {
        readMs = Math.min(readMs, await timed(reading))
        referenceMs = Math.min(referenceMs, await timed(reference))
      }
      if (r >= 1 && readMs / referenceMs > 4 * MAX_RELATIVE) break // far off: no need to repeat it
    }
    expect(readMs / referenceMs, `rows ${readMs.toFixed(0)} ms, tokenizing ${referenceMs.toFixed(0)} ms`).toBeLessThan(MAX_RELATIVE)

    // and the details pass over the same sheet
    const details = await Effect.runPromise(Workbook.sheetDetails(wb, 0))
    expect(details.merges).toEqual([])
  }, 300_000)
})

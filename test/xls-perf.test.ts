/**
 * A 65 536-row XLS (BIFF8's maximum) streams with bounded work between
 * event-loop turns, in time linear in its size. Checked without wall-clock
 * limits: the rows per turn as a count, and linearity as a ratio to
 * reference work in the same process (walking the same records and making
 * one plain cell per record), which machine load slows alike.
 */
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { RecordReader } from "../src/internal/biff/records.ts"
import { open } from "../src/internal/cfb/cfb.ts"
import { BATCH_RECORDS } from "../src/internal/xls/sheet.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { bigXls } from "./helpers/big-xls.ts"
import { countYields } from "./helpers/yields.ts"

const ROWS = 65_536
/**
 * Reading the rows may take at most this many times the reference work.
 * Measured 3.3-6.2x (4 parallel runs, 8 CPU burners); a reader rescanning
 * all earlier rows (65 536² / 2 steps) is hundreds of times above it.
 */
const MAX_RELATIVE = 12
/** Each measurement repeats the work, so a run takes long enough for the clock. */
const REPEAT = 4

const timed = async (run: () => Promise<unknown>): Promise<number> => {
  const start = performance.now()
  // progress: k up to REPEAT
  for (let k = 0; k < REPEAT; k++) await run()
  return performance.now() - start
}

describe("xls: 65 536 rows", () => {
  it("stream with bounded rows between event-loop turns, in linear time", async () => {
    const bytes = bigXls(ROWS)
    const wb = await Effect.runPromise(Xls.read(bytes))
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
    expect(hidden).toBe(Math.floor(ROWS / 10))
    expect(cells).toBe(4 * ROWS)
    // every emitted row took at least one record, and a pull reads at most
    // BATCH_RECORDS records (rows declared in one batch may come out in the next)
    expect(run.maxPerTurn).toBeLessThanOrEqual(2 * BATCH_RECORDS)
    expect(run.yields).toBeGreaterThanOrEqual((5 * ROWS) / BATCH_RECORDS / 2)

    // linear: compared with walking the same records, best of 3, interleaved
    const cfb = open(bytes, SpreadsheetLimits.defaults)
    if ("_tag" in cfb) throw new Error(cfb._tag)
    const stream = cfb.read(cfb.find("Workbook")!)
    if ("_tag" in stream) throw new Error(stream._tag)
    const reference = async () => {
      const reader = new RecordReader(stream, "Workbook")
      const made: Array<Cell.Cell> = []
      // progress: every next() advances the reader
      for (let r = reader.next(); r !== undefined && !("_tag" in r); r = reader.next()) made.push(Cell.number(r.type & 0xFF, r.offset))
      return made.length
    }
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
      if (r >= 1 && readMs / referenceMs > 4 * MAX_RELATIVE) break
    }
    expect(readMs / referenceMs, `rows ${readMs.toFixed(0)} ms, reference ${referenceMs.toFixed(0)} ms`).toBeLessThan(MAX_RELATIVE)
    expect((await Effect.runPromise(Workbook.sheetDetails(wb, 0))).merges).toEqual([])
  }, 300_000)
})

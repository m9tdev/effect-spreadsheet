/**
 * PLAN.md performance check: a 200 000-row XLSX streams with bounded work
 * between event-loop turns. Checked by work (rows per turn), not by the
 * clock; scripts/bench/xlsx-rows.ts measures the actual pauses.
 */
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { bigXlsx, MIN_ROW_BYTES } from "./helpers/big-xlsx.ts"
import { countYields } from "./helpers/yields.ts"

/** The most bytes of sheet XML one pull handles: one inflated or stored piece. */
const PIECE_BYTES = 64 * 1024

describe("xlsx: 200 000 rows", () => {
  it("stream with at most one piece of rows between event-loop turns", async () => {
    const bytes = await bigXlsx(200_000)
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
    expect(rows).toBe(200_000)
    expect(hidden).toBe(20_000)
    expect(cells).toBe(800_000)
    // the work between two turns is at most one piece of XML worth of rows
    expect(run.maxPerTurn).toBeLessThanOrEqual(Math.ceil(PIECE_BYTES / MIN_ROW_BYTES) + 1)
    expect(run.yields).toBeGreaterThanOrEqual(200_000 / (Math.ceil(PIECE_BYTES / MIN_ROW_BYTES) + 1))
    // and the details pass over the same sheet
    const details = await Effect.runPromise(Workbook.sheetDetails(wb, 0))
    expect(details.merges).toEqual([])
  }, 180_000)
})

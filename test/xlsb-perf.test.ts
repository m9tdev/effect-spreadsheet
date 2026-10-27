/**
 * A 200 000-row XLSB streams with bounded work between event-loop turns,
 * in work linear in its size. Checked without wall-clock limits: records
 * and rows per turn as counts; linearity as the reader's own record
 * counter (`work`, exact whatever the load) from n to 16n rows, and as the
 * scaling of its CPU time, with a margin for load.
 */
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { sheetDetails, sheetRows, work } from "../src/internal/xlsb/sheet.ts"
import { SLICE_BYTES } from "../src/internal/xlsb/slices.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsb from "../src/xlsb/index.ts"
import * as B from "./helpers/xlsb.ts"
import { countYields } from "./helpers/yields.ts"

const ROWS = 200_000
const SMALL = ROWS / 16
/** As in xls-perf: 16n rows may take at most this many times 16 x the CPU time of n rows (load only slows a run down). */
const MAX_SCALING = 2.2
const ROUNDS = 8

/**
 * Rows of every kind, every 10th hidden: a shared string, an RK number, a
 * real with a date style, an inline string, every 7th row without cells,
 * every 1024th a row of 300 repeated blanks (sorted at its end).
 */
const bigXlsb = (rows: number, method: "store" | "deflate" = "deflate") => {
  const body: Array<Uint8Array> = []
  // progress: r up to rows
  for (let r = 0; r < rows; r++) {
    body.push(B.rowHdr(r, r % 10 === 9))
    if (r % 7 === 3) continue
    body.push(B.isst(0, r % 3), B.rk(1, B.rkInt(r)), B.real(2, 46000 + (r % 1000), 1), B.cellSt(3, `row ${r}`))
    if (r % 1024 === 512) body.push(...Array.from({ length: 300 }, (_, k) => B.blank(300 - k)))
  }
  return B.xlsb({ strings: ["alpha", "beta", "gamma"], xfs: [0, 14], method, sheets: [{ name: "Big", part: B.sheetBin(body, { after: [B.mergeCell(0, 1, 0, 1)] }) }] })
}

/** CPU time of this process for `repeat` runs, in ms. */
const cpuMs = async (run: () => Promise<unknown>, repeat: number): Promise<number> => {
  const start = process.cpuUsage()
  // progress: k up to repeat
  for (let k = 0; k < repeat; k++) await run()
  const used = process.cpuUsage(start)
  return (used.user + used.system) / 1000
}

describe("xlsb: 200 000 rows", () => {
  for (const method of ["deflate", "store"] as const) {
    it(`stream (${method}) with at most one slice of records between event-loop turns, every hidden row included`, async () => {
      const wb = await Effect.runPromise(Xlsb.read(await bigXlsb(ROWS, method)))
      let rows = 0
      let hidden = 0
      const run = await countYields(
        () =>
          Effect.runPromise(Stream.runForEach(Workbook.rows(wb, 0), (row) =>
            Effect.sync(() => {
              rows++
              if (row.hidden) hidden++
            }))),
        () => work.records
      )
      expect([rows, hidden]).toEqual([ROWS, ROWS / 10])
      // a slice of SLICE_BYTES gives at most SLICE_BYTES / 2 records (and a
      // turn may see the end of one slice's rows and the next slice)
      expect(run.maxPerTurn).toBeLessThanOrEqual(SLICE_BYTES)
      expect(run.yields).toBeGreaterThanOrEqual(ROWS / (SLICE_BYTES / 2))
      const details = await countYields(() => Effect.runPromise(Workbook.sheetDetails(wb, 0)), () => work.records)
      expect(details.value.merges).toHaveLength(1)
      expect(details.maxPerTurn).toBeLessThanOrEqual(SLICE_BYTES)
    }, 300_000)
  }

  it("a part handed over in one piece is still read a slice per turn", async () => {
    const body: Array<Uint8Array> = []
    // progress: r up to ROWS
    for (let r = 0; r < ROWS; r++) body.push(B.rowHdr(r), B.rk(0, B.rkInt(r)))
    const part = B.sheetBin(body)
    const limits = SpreadsheetLimits.defaults
    const context = { part: "one.bin", strings: undefined, styles: { formats: [] }, date1904: false, limits }
    let rows = 0
    const run = await countYields(
      () => Effect.runPromise(Stream.runForEach(sheetRows(Stream.make(part), context), () => Effect.sync(() => void rows++))),
      () => work.records
    )
    expect(rows).toBe(ROWS)
    expect(run.maxPerTurn).toBeLessThanOrEqual(SLICE_BYTES)
    const details = await countYields(() => Effect.runPromise(sheetDetails(Stream.make(part), "one.bin", limits)), () => work.records)
    expect(details.maxPerTurn).toBeLessThanOrEqual(SLICE_BYTES)
  }, 300_000)

  it("read in linear work: 16x the rows take 16x the records", async () => {
    const records = async (rows: number) => {
      const wb = await Effect.runPromise(Xlsb.read(await bigXlsb(rows)))
      const before = work.records
      const out = Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0))))
      expect(out.map((r) => r.index)).toEqual(Array.from({ length: rows }, (_, r) => r))
      return work.records - before
    }
    const small = await records(SMALL)
    const big = await records(ROWS)
    expect(big / (16 * small), `16n rows ${big} records, n rows ${small} records`).toBeLessThan(1.01)
  }, 300_000)

  it("read in linear time: 16x the rows take about 16x the CPU time", async () => {
    const small = await Effect.runPromise(Xlsb.read(await bigXlsb(SMALL)))
    const big = await Effect.runPromise(Xlsb.read(await bigXlsb(ROWS)))
    const reading = (wb: Workbook.Workbook) => () => Effect.runPromise(Stream.runDrain(Workbook.rows(wb, 0)))
    let smallMs = Infinity
    let bigMs = Infinity
    // progress: r up to ROUNDS; best of each size, interleaved, until under the limit
    for (let r = 0; r < ROUNDS; r++) {
      if (r % 2 === 0) {
        smallMs = Math.min(smallMs, await cpuMs(reading(small), 32))
        bigMs = Math.min(bigMs, await cpuMs(reading(big), 2))
      } else {
        bigMs = Math.min(bigMs, await cpuMs(reading(big), 2))
        smallMs = Math.min(smallMs, await cpuMs(reading(small), 32))
      }
      if (r >= 2 && bigMs / smallMs < MAX_SCALING) break
    }
    expect(bigMs / smallMs, `16n rows ${bigMs.toFixed(0)} ms, 16 x n rows ${smallMs.toFixed(0)} ms (CPU)`).toBeLessThan(MAX_SCALING)
  }, 300_000)
})

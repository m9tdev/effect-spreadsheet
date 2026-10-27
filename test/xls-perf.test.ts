/**
 * A 65 536-row XLS (BIFF8's maximum) streams with bounded work between
 * event-loop turns, in work linear in its size. Checked without wall-clock
 * limits: the rows and cells per turn as counts; linearity as the reader's
 * own step counter (`work`, exact whatever the load) from n to 16n rows,
 * and, for work outside its loops, as the scaling of its CPU time, with a
 * margin for load.
 */
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { BATCH_CELLS, BATCH_RECORDS, work } from "../src/internal/xls/sheet.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { bigXls } from "./helpers/big-xls.ts"
import * as X from "./helpers/xls.ts"
import { countYields } from "./helpers/yields.ts"

const ROWS = 65_536
const SMALL = ROWS / 16
/**
 * Reading 16n rows may take at most this many times 16 x the CPU time of n
 * rows. Measured 1.05-1.25 alone and up to 1.8 with 8 runs in parallel
 * plus the full suite; a reader that walks all rows emitted so far every 64
 * rows is 2.6-2.9 alone (about 2 under that load), a quadratic one about 12.
 * Smaller excesses can't be told from load: the step counter checks the
 * reader's own loops exactly.
 */
const MAX_SCALING = 2.2
/** Steps for 16n rows over 16 x the steps for n rows: exact, so only rounding room. */
const MAX_STEP_SCALING = 1.01
/** Rounds of interleaved measurements at most (the best of each size counts). */
const ROUNDS = 8

/** One RK cell per row, no ROW records: the least work per row, so extra work per row shows most. */
const leanXls = (rows: number) => X.xls({ sheets: [{ name: "Lean", records: Array.from({ length: rows }, (_, r) => X.rk(r, 0, 0, X.rkInt(r))) }] })

/**
 * Rows of every kind: ROW records (every 10th hidden) with one cell, every
 * 7th row ROW-only, every 32nd row's cells continued after a later row
 * (merged rows), a few wide rows of repeated columns (compacted).
 */
const mixedXls = (rows: number) => {
  const records: Array<Uint8Array> = []
  // progress: r up to rows
  for (let r = 0; r < rows; r++) {
    records.push(X.row(r, r % 10 === 9))
    if (r % 7 !== 3) records.push(X.rk(r, 1, 0, X.rkInt(r)))
    if (r % 32 === 31) records.push(X.number(r - 16, 0, 0, r))
    if (r % 1024 === 512) records.push(X.mulBlank(r, 10, new Array(300).fill(0)), X.mulBlank(r, 5, new Array(300).fill(0)))
  }
  return X.xls({ sheets: [{ name: "Mixed", records }] })
}

/** CPU time of this process for `repeat` runs, in ms (other processes' load doesn't count, unlike wall time). */
const cpuMs = async (run: () => Promise<unknown>, repeat: number): Promise<number> => {
  const start = process.cpuUsage()
  // progress: k up to repeat
  for (let k = 0; k < repeat; k++) await run()
  const used = process.cpuUsage(start)
  return (used.user + used.system) / 1000
}

describe("xls: 65 536 rows", () => {
  it("stream with bounded rows between event-loop turns", async () => {
    const wb = await Effect.runPromise(Xls.read(bigXls(ROWS)))
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
    // a chunk of rows has at most BATCH_RECORDS rows (and the turn that
    // takes it may follow one that took the previous chunk's last rows)
    expect(run.maxPerTurn).toBeLessThanOrEqual(2 * BATCH_RECORDS)
    expect(run.yields).toBeGreaterThanOrEqual((5 * ROWS) / BATCH_RECORDS / 2)
    expect((await Effect.runPromise(Workbook.sheetDetails(wb, 0))).merges).toEqual([])
  }, 300_000)

  it("stream wide rows (MULBLANK) with bounded cells between event-loop turns", async () => {
    // 64 rows of 16 384 empty cells, 4 MULBLANK records each: few records, many cells
    const xfs = new Array<number>(4096).fill(0)
    const records = Array.from({ length: 64 }, (_, r) => [0, 1, 2, 3].map((k) => X.mulBlank(r, 4096 * k, xfs))).flat()
    // and 48 of them (under the held-cells cap) all waiting for a last cell
    // of row 0 at the end: they complete together, and come out in chunks
    const wb = await Effect.runPromise(Xls.read(X.xls({
      sheets: [{ name: "Wide", records }, { name: "Waiting", records: [...records.slice(0, 4 * 48), X.number(0, 0, 0, 1)] }]
    })))
    for (const [sheet, rows] of [[0, 64], [1, 48]] as const) {
      let cells = 0
      const run = await countYields(
        () => Effect.runPromise(Stream.runForEach(Workbook.rows(wb, sheet), (row) => Effect.sync(() => void (cells += row.cells.length)))),
        () => cells
      )
      expect(cells).toBe(rows * 16_384)
      // a chunk ends once it passes BATCH_CELLS; a row has at most 16 384 cells
      expect(run.maxPerTurn, `sheet ${sheet}`).toBeLessThanOrEqual(BATCH_CELLS + 16_384)
      expect(run.yields).toBeGreaterThanOrEqual((rows * 16_384) / (BATCH_CELLS + 16_384))
    }
    // and so does reading: a batch ends once its cells pass BATCH_CELLS (a
    // MULBLANK adds at most 4 109), a step per cell plus a few per record and row
    const steps = await countYields(() => Effect.runPromise(Stream.runDrain(Workbook.rows(wb, 0))), () => work.steps)
    expect(steps.maxPerTurn).toBeLessThanOrEqual(BATCH_CELLS + 4109 + 4 * BATCH_RECORDS)
  }, 300_000)

  it("sorting unsorted rows counts towards a batch: bounded steps per turn", async () => {
    // rows of 32 768 cells in two passes over 16 384 columns (unsorted, not
    // compacted while read), each sorted when it ends: one sheet with each
    // row closed in turn, one with 30 rows all ending at once (row 0 last);
    // and one row of 16 passes, compacted while read every 16 384 cells
    const xfs = new Array<number>(4096).fill(0)
    const wideRow = (r: number) => [0, 1, 2, 3, 0, 1, 2, 3].map((k) => X.mulBlank(r, 4096 * k, xfs))
    const oneByOne = Array.from({ length: 40 }, (_, r) => wideRow(r)).flat()
    const together = [X.blank(0, 0, 0), ...Array.from({ length: 30 }, (_, r) => wideRow(r + 1)).flat(), X.blank(0, 1, 0)]
    const repeated = Array.from({ length: 8 }, () => wideRow(0)).flat()
    const wb = await Effect.runPromise(Xls.read(X.xls({
      sheets: [{ name: "One by one", records: oneByOne }, { name: "Together", records: together }, { name: "Repeated", records: repeated }]
    })))
    for (const sheet of [0, 1, 2]) {
      const steps = await countYields(() => Effect.runPromise(Stream.runDrain(Workbook.rows(wb, sheet))), () => work.steps)
      // a batch: BATCH_CELLS of reading and sorting, then at most one more
      // record (4 109 cells) or one more row sorted (32 768), plus a few
      // steps per record and row
      expect(steps.maxPerTurn, `sheet ${sheet}`).toBeLessThanOrEqual(BATCH_CELLS + 32_768 + 4109 + 4 * BATCH_RECORDS)
    }
  }, 300_000)

  it("read in linear work: 16x the rows take 16x the reader's steps", async () => {
    const steps = async (rows: number) => {
      const wb = await Effect.runPromise(Xls.read(mixedXls(rows)))
      const before = work.steps
      const out = Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0))))
      expect(out.map((r) => r.index)).toEqual(Array.from({ length: rows }, (_, r) => r))
      return work.steps - before
    }
    const small = await steps(SMALL)
    const big = await steps(ROWS)
    // a few steps per record: scan, read, cell, row, queue
    expect(small / (SMALL * 2.2)).toBeLessThan(8)
    expect(big / (16 * small), `16n rows ${big} steps, n rows ${small} steps`).toBeLessThan(MAX_STEP_SCALING)
  }, 300_000)

  it("read in linear time: 16x the rows take about 16x the CPU time", async () => {
    const small = await Effect.runPromise(Xls.read(leanXls(SMALL)))
    const big = await Effect.runPromise(Xls.read(leanXls(ROWS)))
    const reading = (wb: Workbook.Workbook) => () => Effect.runPromise(Stream.runDrain(Workbook.rows(wb, 0)))
    let smallMs = Infinity
    let bigMs = Infinity
    // progress: r increases up to ROUNDS; best of each size, interleaved,
    // until the ratio is under the limit (load only makes a run slower)
    for (let r = 0; r < ROUNDS; r++) {
      // 16 small runs read as many rows as one big run; 4 of those for a measurable time
      if (r % 2 === 0) {
        smallMs = Math.min(smallMs, await cpuMs(reading(small), 64))
        bigMs = Math.min(bigMs, await cpuMs(reading(big), 4))
      } else {
        bigMs = Math.min(bigMs, await cpuMs(reading(big), 4))
        smallMs = Math.min(smallMs, await cpuMs(reading(small), 64))
      }
      if (r >= 2 && bigMs / smallMs < MAX_SCALING) break
    }
    expect(bigMs / smallMs, `16n rows ${bigMs.toFixed(0)} ms, 16 x n rows ${smallMs.toFixed(0)} ms (CPU)`).toBeLessThan(MAX_SCALING)
  }, 300_000)
})

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
import * as Cell from "../src/Cell.ts"
import { parse } from "../src/internal/xml/parse.ts"
import { open } from "../src/internal/zip/archive.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { bigXlsx, MIN_ROW_BYTES } from "./helpers/big-xlsx.ts"
import { countYields } from "./helpers/yields.ts"
import { xlsx } from "./helpers/xlsx.ts"

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

/**
 * The longest stretch without an event-loop turn while `run` is in flight,
 * in ms: a setImmediate ticker that re-arms itself notes the time between
 * its runs.
 */
const longestGap = async (run: () => Promise<unknown>): Promise<number> => {
  let last = performance.now()
  let gap = 0
  let running = true
  const sample = () => {
    const now = performance.now()
    gap = Math.max(gap, now - last)
    last = now
  }
  const tick = () => {
    if (!running) return
    sample()
    setImmediate(tick)
  }
  setImmediate(tick)
  try {
    await run()
    sample()
    return gap
  } finally {
    running = false
  }
}

/**
 * Styles read before any row is out: classifying number-format codes must
 * not block the event loop for long, whatever the number of cell styles or
 * the length of the codes. Without a per-format bound this work grew with
 * the cell styles times the code length (seconds for a few KB file), or
 * with the total length of the codes in one stretch after the parse.
 * The bound is generous so a loaded machine doesn't fail it: these cases
 * take seconds in one stretch when the work isn't spread out, and a few
 * ms per stretch when it is.
 */
const MAX_GAP_MS = 500
const SERIAL = 45_000 // 2023-03-15

const cellsOf = async (bytes: Uint8Array) => {
  const wb = await Effect.runPromise(Xlsx.read(bytes))
  const rows = Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0))))
  return rows.flatMap((row) => row.cells.map((c) => Cell.value(c) instanceof Date))
}

describe("xlsx: number formats in the styles part", () => {
  it("many cell styles sharing one long custom code don't block the event loop", async () => {
    const XFS = 10_000
    const long = "a".repeat(10_000)
    // xf 0..XFS-1 use the long code; then a custom date, a built-in date and a built-in number
    const xfs = [...Array.from({ length: XFS }, () => 164), 165, 14, 4]
    const bytes = await xlsx({
      sheets: [{
        name: "S",
        data: `<row r="1"><c r="A1" s="0"><v>${SERIAL}</v></c><c r="B1" s="${XFS - 1}"><v>${SERIAL}</v></c>` +
          `<c r="C1" s="${XFS}"><v>${SERIAL}</v></c><c r="D1" s="${XFS + 1}"><v>${SERIAL}</v></c>` +
          `<c r="E1" s="${XFS + 2}"><v>${SERIAL}</v></c></row>`
      }],
      numFmts: { 164: long, 165: "yyyy-mm-dd" },
      xfs
    })
    let dates: Array<boolean> = []
    const gap = await longestGap(async () => {
      dates = await cellsOf(bytes)
    })
    expect(dates).toEqual([false, false, true, true, false])
    expect(gap, `longest gap ${gap.toFixed(0)} ms`).toBeLessThan(MAX_GAP_MS)
  }, 120_000)

  it("many distinct long custom codes don't block the event loop", async () => {
    // 256 codes of 250 000 characters: 64 MB of codes, under the inflate
    // budget of a small file; even ids end in a date letter
    const IDS = 256
    const numFmts: Record<number, string> = {}
    const filler = "a".repeat(250_000 - 1)
    for (let k = 0; k < IDS; k++) numFmts[164 + k] = filler + (k % 2 === 0 ? "d" : "0")
    const xfs = Array.from({ length: IDS }, (_, k) => 164 + k)
    const cells = xfs.map((_, k) => `<c s="${k}"><v>${SERIAL}</v></c>`).join("")
    const bytes = await xlsx({ sheets: [{ name: "S", data: `<row r="1">${cells}</row>` }], numFmts, xfs })
    let dates: Array<boolean> = []
    const gap = await longestGap(async () => {
      dates = await cellsOf(bytes)
    })
    expect(dates).toEqual(xfs.map((_, k) => k % 2 === 0))
    expect(gap, `longest gap ${gap.toFixed(0)} ms`).toBeLessThan(MAX_GAP_MS)
  }, 120_000)

  it("many cell styles with distinct ids don't block the event loop", async () => {
    // 2 500 000 cell styles, each with its own id that has no code (neither
    // custom nor built-in): about 60 MB of styles, under the inflate budget
    // of a small file; then a custom date code and built-in date and number ids.
    // Looking up each style's format after the parse is one stretch of up to
    // a few hundred ms on a loaded machine; building an entry per distinct
    // id takes seconds. The bound sits between the two.
    const XFS = 2_500_000
    const MAX_STYLES_GAP_MS = 800
    const xfs = [...Array.from({ length: XFS }, (_, k) => 1000 + k), 165, 14, 4]
    const bytes = await xlsx({
      sheets: [{
        name: "S",
        data: `<row r="1"><c r="A1" s="0"><v>${SERIAL}</v></c><c r="B1" s="${XFS - 1}"><v>${SERIAL}</v></c>` +
          `<c r="C1" s="${XFS}"><v>${SERIAL}</v></c><c r="D1" s="${XFS + 1}"><v>${SERIAL}</v></c>` +
          `<c r="E1" s="${XFS + 2}"><v>${SERIAL}</v></c></row>`
      }],
      numFmts: { 165: "yyyy-mm-dd" },
      xfs
    })
    let dates: Array<boolean> = []
    const gap = await longestGap(async () => {
      dates = await cellsOf(bytes)
    })
    expect(dates).toEqual([false, false, true, true, false])
    expect(gap, `longest gap ${gap.toFixed(0)} ms`).toBeLessThan(MAX_STYLES_GAP_MS)
  }, 120_000)
})

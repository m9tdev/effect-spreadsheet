/**
 * Custom number formats are classified (date or not) while the styles part
 * is read, a slice per event-loop turn, not when a cell first uses them:
 * a row whose cells use many distinct long codes doesn't block the loop
 * while its rows are streamed. Checked as the longest wall-clock gap
 * between two turns of the loop, with a bound far above what the reader
 * needs (a few ms) and far below what classifying every code in one step
 * takes (over half a second here), so load on the machine doesn't flip it.
 */
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import type * as Cell from "../src/Cell.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsb from "../src/xlsb/index.ts"
import * as B from "./helpers/xlsb.ts"

/** Distinct custom codes, one cell style each. */
const CODES = 120
/** Characters per code: just under the default `maxStringLength`. */
const CODE_LENGTH = 249_999
/** Longest gap between two event-loop turns allowed, in ms. */
const MAX_GAP_MS = 250

/** The longest gap between two turns of the event loop while `run` is in flight, in ms. */
const longestGap = async <A>(run: () => Promise<A>): Promise<{ value: A; gapMs: number }> => {
  let running = true
  let last = performance.now()
  let gapMs = 0
  const tick = () => {
    if (!running) return
    const now = performance.now()
    gapMs = Math.max(gapMs, now - last)
    last = now
    setImmediate(tick)
  }
  setImmediate(tick)
  try {
    const value = await run()
    gapMs = Math.max(gapMs, performance.now() - last)
    return { value, gapMs }
  } finally {
    running = false
  }
}

/**
 * A BrtFmt record for format `id`, written straight into bytes (the
 * record helpers build arrays of numbers, slow at this size): `code` is
 * ASCII.
 */
const fmt = (id: number, code: string): Uint8Array => {
  const size = 2 + 4 + 2 * code.length
  const head = [...B.varint(0x2C), ...B.varint(size), ...B.u16(id), ...B.u32(code.length)]
  const out = new Uint8Array(head.length + 2 * code.length)
  out.set(head)
  // progress: k up to code.length
  for (let k = 0; k < code.length; k++) out[head.length + 2 * k] = code.charCodeAt(k)
  return out
}

/**
 * Style k uses custom format 164 + k: `CODE_LENGTH` characters of `a`, the
 * last one `d` (a date code) for odd k, `0` for even k, so telling them
 * apart scans the whole code. One more style uses built-in format 14.
 * One row has one cell per style.
 */
const manyFormats = () => {
  const xfs = [...Array.from({ length: CODES }, (_, k) => 164 + k), 14]
  const stylesPart = B.concat([
    B.record(0x267, B.u32(CODES)),
    ...Array.from({ length: CODES }, (_, k) => fmt(164 + k, "a".repeat(CODE_LENGTH - 1) + (k % 2 === 1 ? "d" : "0"))),
    B.record(0x268),
    B.record(0x269, B.u32(xfs.length)),
    ...xfs.map(B.xf),
    B.record(0x26A)
  ])
  const cells = xfs.map((_, k) => B.real(k, 1.5, k))
  return B.xlsb({ sheets: [{ name: "S", part: B.sheetBin([B.rowHdr(0), ...cells]) }], stylesPart })
}

describe("xlsb: many long custom number formats", () => {
  it("a row using every style streams without a long event-loop gap, with dates typed as before", async () => {
    const bytes = await manyFormats()
    const wb = await Effect.runPromise(Xlsb.read(bytes))
    const rows = await longestGap(() => Effect.runPromise(Stream.runCollect(Workbook.rows(wb, 0))))
    const cells: ReadonlyArray<Cell.Cell> = Array.from(rows.value)[0]!.cells
    expect(cells.map((c) => c._tag)).toEqual([
      ...Array.from({ length: CODES }, (_, k) => (k % 2 === 1 ? "Date" : "Number")),
      "Date"
    ])
    expect(cells[1]!.numFmt).toHaveLength(CODE_LENGTH)
    expect(cells[CODES]!.numFmt).toBe("mm-dd-yy")
    expect(rows.gapMs, `longest gap while streaming rows: ${rows.gapMs.toFixed(0)} ms`).toBeLessThan(MAX_GAP_MS)
  }, 120_000)
})

/**
 * Page side of `browser-inflate.ts`: bundled twice for the browser, once
 * with the default inflater (`DecompressionStream`) and once with the
 * TypeScript one swapped in at the archive's import of `./inflate.ts` (as
 * `own-inflate.ts` does in Node). Each bundle registers itself under its
 * mode on `globalThis.inflateBench`. Not a test.
 */
import { Effect, Stream } from "effect"
import type * as Cell from "../../src/Cell.ts"
import * as Workbook from "../../src/Workbook.ts"
import { inflateRawWeb } from "../../src/internal/zip/inflate.ts"
import { inflate } from "../../src/internal/zip/rfc1951.ts"
import * as Xlsx from "../../src/xlsx/index.ts"

declare const BENCH_MODE: string

/** What one read of the workbook reports. */
export interface ReadResult {
  readonly rows: number
  /** FNV-1a of every row's index, hidden flag and cells, or 0 when not asked. */
  readonly hash: number
  readonly ms: number
  /** Longest gap between ticks of a 1 ms interval: the longest main-thread pause seen. */
  readonly maxGapMs: number
  /** Longest `longtask` entry, where the browser reports them (Chromium), else -1. */
  readonly maxLongTaskMs: number
}

const fnv = (hash: number, text: string): number => {
  let h = hash
  // progress: i up to text.length
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i)
    h = Math.imul(h, 16777619)
  }
  return h >>> 0
}

const cellText = (cell: Cell.Cell): string => {
  const value = "value" in cell ? (cell.value instanceof Date ? `D${cell.value.getTime()}` : String(cell.value)) : ""
  return `${cell._tag}|${cell.column}|${value}|${cell.numFmt ?? ""}|${cell.formula ?? ""};`
}

const watchPauses = () => {
  let maxGap = 0
  let last = performance.now()
  const timer = setInterval(() => {
    const now = performance.now()
    maxGap = Math.max(maxGap, now - last)
    last = now
  }, 1)
  let maxLongTask = -1
  let observer: PerformanceObserver | undefined
  if (PerformanceObserver.supportedEntryTypes.includes("longtask")) {
    maxLongTask = 0
    observer = new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) maxLongTask = Math.max(maxLongTask, entry.duration)
    })
    observer.observe({ type: "longtask" })
  }
  return async () => {
    await new Promise((resolve) => setTimeout(resolve, 20))
    clearInterval(timer)
    if (observer !== undefined) {
      for (const entry of observer.takeRecords()) maxLongTask = Math.max(maxLongTask, entry.duration)
      observer.disconnect()
    }
    return { maxGapMs: Math.max(maxGap, performance.now() - last), maxLongTaskMs: maxLongTask }
  }
}

/** Reads every row of the first sheet; with `hash`, also hashes their contents (slower: for the equality check). */
const readRows = async (bytes: Uint8Array, hash: boolean): Promise<ReadResult> => {
  const stop = watchPauses()
  let rows = 0
  let h = 2166136261
  const start = performance.now()
  const wb = await Effect.runPromise(Xlsx.read(bytes))
  await Effect.runPromise(
    Stream.runForEach(Workbook.rows(wb, 0), (row) =>
      Effect.sync(() => {
        rows++
        if (!hash) return
        let text = `${row.index}|${row.hidden ? 1 : 0}:`
        for (const cell of row.cells) text += cellText(cell)
        h = fnv(h, text)
      }))
  )
  const ms = performance.now() - start
  const pauses = await stop()
  return { rows, hash: hash ? h : 0, ms, ...pauses }
}

/**
 * Raw inflate of one deflate stream: the TypeScript inflate in 64 KiB steps,
 * or `DecompressionStream` through the library's default inflater. Returns
 * the output size, the number of chunks it came in, and the time.
 */
const rawInflate = async (
  data: Uint8Array,
  which: "ts" | "web"
): Promise<{ readonly bytes: number; readonly chunks: number; readonly ms: number }> => {
  const start = performance.now()
  let n = 0
  let chunks = 0
  if (which === "ts") {
    const decoder = inflate(data)
    // progress: each step writes, consumes input or ends
    for (;;) {
      const out = new Uint8Array(65536)
      const step = decoder.next(out, 65536)
      if (step._tag === "Output") {
        n += step.written
        chunks++
      } else if (step._tag === "Done") break
      else if (step._tag === "Failed") throw new Error(step.reason)
    }
  } else {
    const inflater = inflateRawWeb(data)
    // progress: each call returns a chunk or ends
    for (;;) {
      const chunk = await inflater.next()
      if (chunk === undefined) break
      n += chunk.length
      chunks++
    }
  }
  return { bytes: n, chunks, ms: performance.now() - start }
}

const target = globalThis as unknown as { inflateBench?: Record<string, unknown> }
target.inflateBench ??= {}
target.inflateBench[BENCH_MODE] = { readRows, rawInflate }

/**
 * Streams the rows of a synthetic 65 536-row XLS (BIFF8's maximum) and reports the time and
 * the longest event-loop pause (PLAN.md performance check). Not a test: the
 * numbers depend on the machine and its load (cpuMs less so than ms).
 * Run: node scripts/bench/xls-rows.ts [rows] [wide | row-only]; with `wide`, each row is
 * 16 384 empty cells in 4 MULBLANK records instead (few records, many cells);
 * with `row-only`, `rows` ROW records without cells (default 10 million, rows
 * 0 to 65 535 over and over: one record of 20 bytes each).
 */
import { Effect, Stream } from "effect"
import { monitorEventLoopDelay } from "node:perf_hooks"
import * as Workbook from "../../src/Workbook.ts"
import * as Xls from "../../src/xls/index.ts"
import { bigXls } from "../../test/helpers/big-xls.ts"
import * as X from "../../test/helpers/xls.ts"

const wide = process.argv[3] === "wide"
const rowOnly = process.argv[3] === "row-only"
const count = rowOnly ? Number(process.argv[2] ?? 10_000_000) : Math.min(65_536, Number(process.argv[2] ?? 65_536))
const blanks = new Array<number>(4096).fill(0)
/** `n` ROW records in one blob (rows 0 to 65 535, over and over). */
const rowRecords = (n: number) => {
  const size = X.row(0).length
  const out = new Uint8Array(n * size)
  // progress: r up to n
  for (let r = 0; r < n; r++) out.set(X.row(r % 65_536), r * size)
  return out
}
const bytes = wide
  ? X.xls({ sheets: [{ name: "Wide", records: Array.from({ length: count }, (_, r) => [0, 1, 2, 3].map((k) => X.mulBlank(r, 4096 * k, blanks))).flat() }] })
  : rowOnly
  ? X.xls({ sheets: [{ name: "Row only", records: [rowRecords(count)] }] })
  : bigXls(count)
let cells = 0
const histogram = monitorEventLoopDelay({ resolution: 1 })
let rows = 0
let maxGap = 0
let last = performance.now()
const timer = setInterval(() => {
  const now = performance.now()
  maxGap = Math.max(maxGap, now - last)
  last = now
}, 1)
histogram.enable()
const start = performance.now()
const cpuStart = process.cpuUsage()
const wb = await Effect.runPromise(Xls.read(bytes))
// the pause while opening the workbook (CFB, globals), apart from the rows'
await new Promise((r) => setTimeout(r, 5))
const readPause = maxGap
maxGap = 0
await Effect.runPromise(Stream.runForEach(Workbook.rows(wb, 0), (row) =>
  Effect.sync(() => {
    rows++
    cells += row.cells.length
  })))
const elapsed = performance.now() - start
const cpu = process.cpuUsage(cpuStart)
await new Promise((r) => setTimeout(r, 20))
histogram.disable()
clearInterval(timer)
console.log(JSON.stringify({
  rows,
  cells,
  fileBytes: bytes.length,
  ms: Math.round(elapsed),
  // CPU time of this process: less dependent on other load than ms
  cpuMs: Math.round((cpu.user + cpu.system) / 1000),
  rowsPerSecond: Math.round(rows / (elapsed / 1000)),
  readMaxPauseMs: Math.round(readPause),
  maxEventLoopPauseMs: Math.round(maxGap),
  eventLoopDelayP99Ms: Math.round(histogram.percentile(99) / 1e6),
  eventLoopDelayMaxMs: Math.round(histogram.max / 1e6)
}))

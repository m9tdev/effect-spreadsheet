/**
 * Streams the rows of a synthetic 65 536-row XLS (BIFF8's maximum) and reports the time and
 * the longest event-loop pause (PLAN.md performance check). Not a test: the
 * numbers depend on the machine and its load (cpuMs less so than ms).
 * Run: node scripts/bench/xls-rows.ts [rows] [wide]; with `wide`, each row is
 * 16 384 empty cells in 4 MULBLANK records instead (few records, many cells).
 */
import { Effect, Stream } from "effect"
import { monitorEventLoopDelay } from "node:perf_hooks"
import * as Workbook from "../../src/Workbook.ts"
import * as Xls from "../../src/xls/index.ts"
import { bigXls } from "../../test/helpers/big-xls.ts"
import * as X from "../../test/helpers/xls.ts"

const count = Math.min(65_536, Number(process.argv[2] ?? 65_536))
const wide = process.argv[3] === "wide"
const blanks = new Array<number>(4096).fill(0)
const bytes = wide
  ? X.xls({ sheets: [{ name: "Wide", records: Array.from({ length: count }, (_, r) => [0, 1, 2, 3].map((k) => X.mulBlank(r, 4096 * k, blanks))).flat() }] })
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
  maxEventLoopPauseMs: Math.round(maxGap),
  eventLoopDelayP99Ms: Math.round(histogram.percentile(99) / 1e6),
  eventLoopDelayMaxMs: Math.round(histogram.max / 1e6)
}))

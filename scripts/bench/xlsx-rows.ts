/**
 * Streams the rows of a synthetic 200 000-row XLSX and reports the time and
 * the longest event-loop pause (PLAN.md performance check). Not a test: the
 * numbers depend on the machine and its load (cpuMs less so than ms). Run: node scripts/bench/xlsx-rows.ts [rows]
 */
import { Effect, Stream } from "effect"
import { monitorEventLoopDelay } from "node:perf_hooks"
import * as Workbook from "../../src/Workbook.ts"
import * as Xlsx from "../../src/xlsx/index.ts"
import { bigXlsx } from "../../test/helpers/big-xlsx.ts"

const count = Number(process.argv[2] ?? 200_000)
const bytes = await bigXlsx(count)
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
const wb = await Effect.runPromise(Xlsx.read(bytes))
await Effect.runPromise(Stream.runForEach(Workbook.rows(wb, 0), () => Effect.sync(() => void rows++)))
const elapsed = performance.now() - start
const cpu = process.cpuUsage(cpuStart)
await new Promise((r) => setTimeout(r, 20))
histogram.disable()
clearInterval(timer)
console.log(JSON.stringify({
  rows,
  fileBytes: bytes.length,
  ms: Math.round(elapsed),
  // CPU time of this process: less dependent on other load than ms
  cpuMs: Math.round((cpu.user + cpu.system) / 1000),
  rowsPerSecond: Math.round(rows / (elapsed / 1000)),
  maxEventLoopPauseMs: Math.round(maxGap),
  eventLoopDelayP99Ms: Math.round(histogram.percentile(99) / 1e6),
  eventLoopDelayMaxMs: Math.round(histogram.max / 1e6)
}))

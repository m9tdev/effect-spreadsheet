/**
 * Streams the rows of a synthetic 200 000-row CSV (UTF-8 text with quoted
 * fields, numbers and non-ASCII characters; the delimiter guessed and the
 * encoding detected) and reports the time and the longest event-loop pause.
 * Not a test: the numbers depend on the machine and its load (cpuMs less
 * so than ms). Run: node scripts/bench/csv-rows.ts [rows] [encoding]
 */
import { Effect, Stream } from "effect"
import { monitorEventLoopDelay } from "node:perf_hooks"
import * as Csv from "../../src/csv/index.ts"

const count = Number(process.argv[2] ?? 200_000)
const encoding = process.argv[3] ?? "auto"
const lines: Array<string> = ["id,name,city,amount,note,date"]
// progress: i increases up to count
for (let i = 0; i < count; i++) {
  lines.push(`${i},"Name ${i % 997}, Jr.",Zürich ${i % 31},${(i * 1.25).toFixed(2)},"said ""hi"" ${i % 7}",2024-0${1 + i % 9}-1${i % 10}`)
}
const bytes = new TextEncoder().encode(lines.join("\r\n") + "\r\n")
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
await Effect.runPromise(Stream.runForEach(Csv.rows(bytes, { encoding }), () => Effect.sync(() => void rows++)))
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

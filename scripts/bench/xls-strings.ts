/**
 * Streams the rows of a synthetic XLS whose cells are all formulas with a
 * text result, each followed by its STRING record; the first cell of every
 * 64th row has a 2 000-character result continued over a CONTINUE record.
 * Reports the time and the CPU time. Not a test: the numbers depend on the
 * machine and its load (cpuMs less so than ms).
 * Run: node scripts/bench/xls-strings.ts [rows]   (default 65 536, BIFF8's maximum)
 */
import { Effect, Stream } from "effect"
import * as Workbook from "../../src/Workbook.ts"
import * as Xls from "../../src/xls/index.ts"
import * as X from "../../test/helpers/xls.ts"

const count = Math.min(65_536, Number(process.argv[2] ?? 65_536))
const records: Array<Uint8Array> = []
// progress: r up to count, c up to 4
for (let r = 0; r < count; r++) {
  for (let c = 0; c < 4; c++) {
    records.push(X.formula(r, c, 0, { kind: 0 }))
    if (r % 64 === 0 && c === 0) {
      // 1 000 compressed characters, then 1 000 16-bit ones after the CONTINUE's option byte
      records.push(
        X.record(0x0207, [...X.u16(2000), 0, ...new Array(1000).fill(0x61)]),
        X.record(0x003C, [1, ...new Array(1000).fill(0).flatMap(() => [0x62, 0])])
      )
    } else records.push(X.string(`s${r}-${c}-text`))
  }
}
const bytes = X.xls({ sheets: [{ name: "Strings", records }] })
const start = performance.now()
const cpuStart = process.cpuUsage()
const wb = await Effect.runPromise(Xls.read(bytes))
let rows = 0
await Effect.runPromise(Stream.runForEach(Workbook.rows(wb, 0), () =>
  Effect.sync(() => {
    rows++
  })))
const cpu = process.cpuUsage(cpuStart)
console.log(JSON.stringify({ rows, fileBytes: bytes.length, ms: Math.round(performance.now() - start), cpuMs: Math.round((cpu.user + cpu.system) / 1000) }))

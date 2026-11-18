/**
 * Reads a synthetic XLSB (every cell record kind, long and short forms)
 * several times and reports the median CPU time of opening it and streaming
 * its rows, and of its sheet details. Not a test: the numbers depend on the
 * machine and its load. Stored entries by default, so the time is the
 * record reader's rather than the inflater's.
 * Run: node scripts/bench/xlsb-rows.ts [rows] [rounds] [store|deflate]
 */
import { Effect, Stream } from "effect"
import * as Workbook from "../../src/Workbook.ts"
import * as Xlsb from "../../src/xlsb/index.ts"
import * as B from "../../test/helpers/xlsb.ts"

const count = Number(process.argv[2] ?? 200_000)
const rounds = Number(process.argv[3] ?? 9)
const method = process.argv[4] === "deflate" ? "deflate" : "store"

const body: Array<Uint8Array> = []
// progress: r up to count
for (let r = 0; r < count; r++) {
  body.push(B.rowHdr(r, r % 10 === 9))
  body.push(
    B.isst(0, r % 3),
    B.rk(1, B.rkInt(r)),
    B.real(2, 46000 + (r % 1000), 1),
    B.cellSt(3, `row ${r}`),
    B.cellBool(4, r % 2 === 0),
    B.cellError(5, 0x07),
    B.fmlaNum(6, r / 3),
    B.shortRk(B.rkDouble(r + 0.5)),
    B.shortReal(r * 1.25),
    B.shortSt("short"),
    B.rstring(10, "rich"),
    B.blank(11)
  )
}
const bytes = await B.xlsb({
  strings: ["alpha", "beta", "gamma"],
  xfs: [0, 14],
  method,
  sheets: [{ name: "Big", part: B.sheetBin(body, { before: [B.wsDim(0, count, 0, 11)], after: [B.mergeCell(0, 1, 0, 1)] }) }]
})

const cpu = async (run: () => Promise<unknown>): Promise<number> => {
  const start = process.cpuUsage()
  await run()
  const used = process.cpuUsage(start)
  return (used.user + used.system) / 1000
}
const median = (xs: Array<number>) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)]!

let cells = 0
const rowsTimes: Array<number> = []
const detailsTimes: Array<number> = []
// progress: k up to rounds
for (let k = 0; k < rounds; k++) {
  cells = 0
  rowsTimes.push(await cpu(async () => {
    const wb = await Effect.runPromise(Xlsb.read(bytes))
    await Effect.runPromise(Stream.runForEach(Workbook.rows(wb, 0), (row) => Effect.sync(() => void (cells += row.cells.length))))
  }))
  detailsTimes.push(await cpu(async () => {
    const wb = await Effect.runPromise(Xlsb.read(bytes))
    await Effect.runPromise(Workbook.sheetDetails(wb, 0))
  }))
}
console.log(JSON.stringify({
  rows: count,
  cells,
  method,
  fileBytes: bytes.length,
  rowsCpuMsMedian: Math.round(median(rowsTimes)),
  detailsCpuMsMedian: Math.round(median(detailsTimes)),
  rowsCpuMs: rowsTimes.map(Math.round)
}))

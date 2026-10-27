/**
 * Opens synthetic XLS workbooks whose globals substream is nearly all of
 * the file and reports the time and the longest event-loop pause of
 * `Xls.read`. Not a test: the numbers depend on the machine and its load.
 * Run: node scripts/bench/xls-globals.ts <case> [megabytes], where case is
 * `xf` (XF records), `format` (FORMAT records), `sst` (one SST of short
 * strings over CONTINUEs), `names` (NAME records) or `continue` (empty
 * CONTINUEs after an SST) or `xf-formats` (4 096 FORMATs with 8 000-character
 * codes and one XF per id, so the XF pass parses every code; megabytes is
 * ignored). Limits that would stop the flood are lifted.
 */
import { Effect, Exit } from "effect"
import { monitorEventLoopDelay } from "node:perf_hooks"
import * as Xls from "../../src/xls/index.ts"

const kind = process.argv[2] ?? "xf"
const target = Number(process.argv[3] ?? 100) * 1024 * 1024

/** A growable little-endian record writer. */
class Writer {
  bytes = new Uint8Array(target + (1 << 20))
  at = 0
  ensure(n: number) {
    if (this.at + n <= this.bytes.length) return
    const grown = new Uint8Array(Math.max(this.bytes.length * 2, this.at + n))
    grown.set(this.bytes.subarray(0, this.at))
    this.bytes = grown
  }
  u8(n: number) {
    this.ensure(1)
    this.bytes[this.at++] = n & 0xFF
  }
  u16(n: number) {
    this.u8(n)
    this.u8(n >> 8)
  }
  u32(n: number) {
    this.u16(n & 0xFFFF)
    this.u16(n >>> 16)
  }
  record(type: number, body: (w: Writer) => void) {
    this.u16(type)
    const lengthAt = this.at
    this.u16(0)
    body(this)
    const length = this.at - lengthAt - 2
    this.bytes[lengthAt] = length & 0xFF
    this.bytes[lengthAt + 1] = length >> 8
  }
}

const w = new Writer()
w.record(0x0809, (r) => {
  ;[0x0600, 0x0005, 0x0DBB, 1996].forEach((n) => r.u16(n))
  r.u32(0)
  r.u32(6)
})
const name = (k: number) => `N${k.toString(36).padStart(7, "0")}`
let count = 0
if (kind === "sst") {
  // strings of 12 characters: 15 bytes each, 500 per 7 500-byte segment
  const strings = Math.floor(target / 15)
  w.record(0x00FC, (r) => {
    r.u32(strings)
    r.u32(strings)
  })
  for (let k = 0; k < strings; k += 500) {
    w.record(0x003C, (r) => {
      for (let j = k; j < Math.min(strings, k + 500); j++) {
        const s = `S${j.toString(36).padStart(11, "0")}`
        r.u16(s.length)
        r.u8(0)
        for (let c = 0; c < s.length; c++) r.u8(s.charCodeAt(c))
      }
    })
  }
  count = strings
} else if (kind === "xf-formats") {
  const code = "0".repeat(8000)
  for (let k = 0; k < 4096; k++) {
    w.record(0x041E, (r) => {
      r.u16(164 + k)
      r.u16(code.length)
      r.u8(0)
      for (let c = 0; c < code.length; c++) r.u8(0x30)
    })
  }
  for (let k = 0; k < 4096; k++) w.record(0x00E0, (r) => [0, 164 + k, 0, 0, 0, 0, 0, 0, 0, 0].forEach((n) => r.u16(n)))
  count = 8192
} else {
  if (kind === "continue") w.record(0x00FC, (r) => (r.u32(0), r.u32(0)))
  // progress: count grows, and the writer by at least 4 bytes per record
  while (w.at < target) {
    if (kind === "xf") w.record(0x00E0, (r) => [0, 14, 0, 0, 0, 0, 0, 0, 0, 0].forEach((n) => r.u16(n)))
    else if (kind === "format") {
      w.record(0x041E, (r) => {
        r.u16(164 + (count % 200))
        const code = "0.00;[Red]0.00"
        r.u16(code.length)
        r.u8(0)
        for (let c = 0; c < code.length; c++) r.u8(code.charCodeAt(c))
      })
    } else if (kind === "names") {
      w.record(0x0018, (r) => {
        const n = name(count)
        // grbit, chKey, cch, cce (7: a PtgRef3d), ixals, itab, four menu/help lengths, flags (8-bit characters)
        r.u16(0)
        r.u8(0)
        r.u8(n.length)
        r.u16(7)
        r.u16(0)
        r.u16(0)
        r.u32(0)
        r.u8(0)
        for (let c = 0; c < n.length; c++) r.u8(n.charCodeAt(c))
        ;[0x3A, 0, 0, 0, 0, 0, 0].forEach((b) => r.u8(b))
      })
    } else if (kind === "continue") w.record(0x003C, () => undefined)
    else throw new Error(`unknown case ${kind}`)
    count++
  }
}
w.record(0x000A, () => {})
const bytes = w.bytes.slice(0, w.at)

const histogram = monitorEventLoopDelay({ resolution: 1 })
let maxGap = 0
let last = performance.now()
const timer = setInterval(() => {
  const now = performance.now()
  maxGap = Math.max(maxGap, now - last)
  last = now
}, 1)
histogram.enable()
const start = performance.now()
const exit = await Effect.runPromise(Effect.exit(Xls.read(bytes, {
  limits: { maxSharedStrings: Infinity, maxDecompressedBytes: Infinity, maxDefinedNames: Infinity }
})))
const elapsed = performance.now() - start
await new Promise((r) => setTimeout(r, 20))
histogram.disable()
clearInterval(timer)
console.log(JSON.stringify({
  case: kind,
  records: count,
  fileBytes: bytes.length,
  outcome: Exit.isSuccess(exit)
    ? { sheets: exit.value.sheets.length, definedNames: exit.value.definedNames.length }
    : String(exit.cause),
  ms: Math.round(elapsed),
  maxEventLoopPauseMs: Math.round(maxGap),
  eventLoopDelayMaxMs: Math.round(histogram.max / 1e6)
}))

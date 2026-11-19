/**
 * Raw inflate throughput: the TypeScript inflate (64 KiB steps, as the
 * archive drives it) against `DecompressionStream` (through the library's
 * default inflater) and node:zlib (synchronous, one call), on large
 * entries, in alternating runs. Reports the median MB/s of output, wall
 * clock. Not a test: the numbers depend on the machine and its load.
 * Run: node scripts/bench/inflate-raw.ts [rounds] [extra files...]
 * Each extra file is a zip; its largest deflated entry is added.
 */
import { readFileSync } from "node:fs"
import * as zlib from "node:zlib"
import * as Result from "effect/Result"
import { locateEntry } from "../../src/internal/zip/archive.ts"
import { readDirectory } from "../../src/internal/zip/directory.ts"
import { inflateRawWeb } from "../../src/internal/zip/inflate.ts"
import { inflate } from "../../src/internal/zip/rfc1951.ts"
import { defaults } from "../../src/SpreadsheetLimits.ts"

const rounds = Number(process.argv[2] ?? 7)
const extra = process.argv.slice(3)

const rows = (count: number) => {
  const parts: Array<string> = []
  for (let i = 1; i <= count; i++) {
    parts.push(
      `<row r="${i}"><c r="A${i}" t="s"><v>${i % 3}</v></c><c r="B${i}"><v>${i * 1.5}</v></c>` +
        `<c r="C${i}" s="1"><v>${46000 + (i % 1000)}</v></c><c r="D${i}" t="inlineStr"><is><t>row ${i}</t></is></c></row>`
    )
  }
  return Buffer.from(parts.join(""))
}
const noise = (n: number) => {
  const out = Buffer.alloc(n)
  let x = 1
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0
    out[i] = (x >>> 16) & 0x3F
  }
  return out
}

const inputs: Array<readonly [string, Uint8Array]> = [
  ["200k-row sheet XML, level 6", zlib.deflateRawSync(rows(200_000))],
  ["200k-row sheet XML, level 1", zlib.deflateRawSync(rows(200_000), { level: 1 })],
  ["200k-row sheet XML, fixed codes", zlib.deflateRawSync(rows(200_000), { strategy: zlib.constants.Z_FIXED })],
  ["16 MB with long repeats (6-bit pattern, period 4 MiB)", zlib.deflateRawSync(noise(16_000_000))],
  ["32 MB stored", zlib.deflateRawSync(noise(32_000_000), { level: 0 })]
]
for (const path of extra) {
  const bytes = new Uint8Array(readFileSync(path))
  const directory = readDirectory(bytes, defaults)
  if (Result.isFailure(directory)) continue
  let best: Uint8Array | undefined
  let name = ""
  for (const entry of directory.success.entries) {
    if (entry.method !== 8) continue
    const located = locateEntry(bytes, directory.success, entry)
    if (Result.isFailure(located)) continue
    const data = bytes.subarray(located.success.data.start, located.success.data.end)
    if (best === undefined || data.length > best.length) {
      best = data
      name = entry.name
    }
  }
  if (best !== undefined) inputs.push([`${path.split("/").pop()}!${name}`, best])
}

const ours = (data: Uint8Array) => {
  const decoder = inflate(data)
  let n = 0
  // progress: each step writes, consumes input or ends
  for (;;) {
    const out = new Uint8Array(65536)
    const step = decoder.next(out, 65536)
    if (step._tag === "Output") n += step.written
    else if (step._tag === "Done") return n
    else if (step._tag === "Failed") throw new Error(step.reason)
  }
}
const platform = async (data: Uint8Array) => {
  const inflater = inflateRawWeb(data)
  let n = 0
  // progress: each call returns a chunk or ends
  for (;;) {
    const chunk = await inflater.next()
    if (chunk === undefined) return n
    n += chunk.length
  }
}
const node = (data: Uint8Array) => zlib.inflateRawSync(data).length

const median = (xs: Array<number>) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)]!
const runners = { ts: ours, DecompressionStream: platform, "node:zlib": node } as const
for (const [label, data] of inputs) {
  const times: Record<string, Array<number>> = { ts: [], DecompressionStream: [], "node:zlib": [] }
  let size = 0
  // progress: k up to rounds
  for (let k = 0; k < rounds; k++) {
    for (const [name, f] of Object.entries(runners)) {
      const start = performance.now()
      size = await f(data)
      times[name]!.push(performance.now() - start)
    }
  }
  const mbps = (name: string) => Math.round(size / 1e6 / (median(times[name]!) / 1000))
  console.log(JSON.stringify({
    input: label,
    compressedBytes: data.length,
    outputBytes: size,
    tsMBps: mbps("ts"),
    decompressionStreamMBps: mbps("DecompressionStream"),
    zlibSyncMBps: mbps("node:zlib"),
    tsVsDecompressionStream: `${Math.round((median(times["ts"]!) / median(times["DecompressionStream"]!) - 1) * 100)}%`
  }))
}

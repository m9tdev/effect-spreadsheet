/**
 * The TypeScript inflate against `DecompressionStream` in real browsers
 * (Chromium, Firefox, WebKit through Playwright): reading every row of the
 * synthetic 200 000-row XLSX (as `xlsx-rows.ts`), and raw inflate MB/s on
 * the inputs of `inflate-raw.ts`. Modes alternate run by run; medians and
 * ranges are printed. Also checks that both modes read the same rows (count
 * and a hash of every row) in each browser. Not a test: the numbers depend
 * on the machine and its load.
 *
 * Playwright and esbuild are not dependencies of this repo: install them
 * anywhere and point `BENCH_TOOLS` at that folder, e.g.
 *   npm i --prefix /tmp/bench-tools playwright esbuild
 *   npx --prefix /tmp/bench-tools playwright install chromium firefox webkit
 *   BENCH_TOOLS=/tmp/bench-tools node scripts/bench/browser-inflate.ts [browsers] [rounds] [rawRounds]
 * browsers: comma-separated, default chromium,firefox,webkit.
 * CPU is the user+system time of the browser's processes (Linux /proc only).
 *
 * Two knobs, to see where the time goes:
 * - BENCH_TS_CHUNK=<bytes> adds a third mode, the TypeScript inflater with
 *   that output chunk size instead of its own.
 * - BENCH_SET_IMMEDIATE=1 gives the page a `setImmediate` (over
 *   `MessageChannel`) before the library loads. Effect yields through
 *   `setImmediate` when there is one, else through `setTimeout(0)`, which
 *   browsers clamp to 4 ms once timeouts nest; the archive yields once per
 *   chunk, so without it the wall time grows with the chunk count.
 */
import { readdirSync, readFileSync } from "node:fs"
import { createServer } from "node:http"
import { createRequire } from "node:module"
import { loadavg } from "node:os"
import { join } from "node:path"
import * as zlib from "node:zlib"
import { bigXlsx } from "../../test/helpers/big-xlsx.ts"

const browsers = (process.argv[2] ?? "chromium,firefox,webkit").split(",")
const rounds = Number(process.argv[3] ?? 10)
const rawRounds = Number(process.argv[4] ?? 8)

// ---- tools, loaded from BENCH_TOOLS (only what is used is typed) ----
interface EsbuildPlugin {
  readonly name: string
  readonly setup: (build: {
    readonly onResolve: (
      options: { readonly filter: RegExp },
      callback: (args: { readonly path: string; readonly importer: string }) => { readonly path: string } | undefined
    ) => void
    readonly onLoad: (
      options: { readonly filter: RegExp },
      callback: (args: { readonly path: string }) => { readonly contents: string; readonly loader: "ts" }
    ) => void
  }) => void
}
interface Esbuild {
  readonly build: (options: Record<string, unknown> & { readonly plugins?: ReadonlyArray<EsbuildPlugin> }) => Promise<{
    readonly outputFiles: ReadonlyArray<{ readonly text: string }>
  }>
}
interface Page {
  readonly goto: (url: string) => Promise<unknown>
  readonly evaluate: <R, A>(f: (arg: A) => R | Promise<R>, arg: A) => Promise<R>
  readonly waitForFunction: (expression: string) => Promise<unknown>
}
interface Browser {
  readonly version: () => string
  readonly newPage: () => Promise<Page>
  readonly close: () => Promise<void>
}
interface BrowserType {
  readonly launch: () => Promise<Browser>
}
const tools = createRequire(join(process.env["BENCH_TOOLS"] ?? process.cwd(), "package.json"))
const esbuild = tools("esbuild") as Esbuild
const playwright = tools("playwright") as Record<string, BrowserType>

// ---- inputs ----
const rowsXml = (count: number) => {
  const parts: Array<string> = []
  // progress: i up to count
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
  // progress: i up to n
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0
    out[i] = (x >>> 16) & 0x3F
  }
  return out
}
const rawInputs: Array<readonly [string, Uint8Array]> = [
  ["200k-row sheet XML, level 6", zlib.deflateRawSync(rowsXml(200_000))],
  ["200k-row sheet XML, level 1", zlib.deflateRawSync(rowsXml(200_000), { level: 1 })],
  ["200k-row sheet XML, fixed codes", zlib.deflateRawSync(rowsXml(200_000), { strategy: zlib.constants.Z_FIXED })],
  ["16 MB with long repeats (6-bit pattern, period 4 MiB)", zlib.deflateRawSync(noise(16_000_000))],
  ["32 MB stored", zlib.deflateRawSync(noise(32_000_000), { level: 0 })]
]
const xlsx = await bigXlsx(200_000)

// ---- bundles: the default inflater, and the TypeScript one at the archive's import ----
const page = new URL("./browser-inflate-page.ts", import.meta.url).pathname
const archive = new URL("../../src/internal/zip/archive.ts", import.meta.url).pathname
const swapped = new URL("./own-inflate-default.ts", import.meta.url).pathname
const tsChunk = process.env["BENCH_TS_CHUNK"]
const modes = ["default", "ts", ...(tsChunk === undefined ? [] : [`ts-${tsChunk}`])]
const swap: EsbuildPlugin = {
  name: "own-inflate",
  setup: (build) =>
    build.onResolve({ filter: /^\.\/inflate\.ts$/ }, (args) => (args.importer === archive ? { path: swapped } : undefined))
}
const chunkSize = (bytes: string): EsbuildPlugin => ({
  name: "own-inflate-chunk",
  setup: (build) =>
    build.onLoad({ filter: /rfc1951\.ts$/ }, (args) => {
      const source = readFileSync(args.path, "utf8")
      const contents = source.replace("const CHUNK_BYTES = 16 * 1024", `const CHUNK_BYTES = ${Number(bytes)}`)
      if (contents === source) throw new Error("CHUNK_BYTES not found in rfc1951.ts")
      return { contents, loader: "ts" }
    })
})
const bundle = async (mode: string) => {
  const result = await esbuild.build({
    entryPoints: [page],
    bundle: true,
    write: false,
    format: "iife",
    platform: "browser",
    target: "es2022",
    define: { BENCH_MODE: JSON.stringify(mode) },
    plugins: mode === "default" ? [] : mode === "ts" ? [swap] : [swap, chunkSize(mode.slice(3))]
  })
  return result.outputFiles[0]!.text
}
const setImmediateShim = process.env["BENCH_SET_IMMEDIATE"] === "1"
  ? `<script>{const q=[];const c=new MessageChannel();c.port1.onmessage=()=>{const f=q.shift();f&&f()};` +
    `globalThis.setImmediate=(f)=>{q.push(f);c.port2.postMessage(0);return f};` +
    `globalThis.clearImmediate=(f)=>{const i=q.indexOf(f);if(i>=0)q[i]=()=>{}}}</script>`
  : ""
const files = new Map<string, { readonly type: string; readonly body: string | Uint8Array }>([
  ["/", {
    type: "text/html",
    body: `<!doctype html><meta charset="utf-8">${setImmediateShim}` +
      modes.map((mode) => `<script src="/${mode}.js"></script>`).join("")
  }],
  ["/big.xlsx", { type: "application/octet-stream", body: xlsx }]
])
for (const mode of modes) files.set(`/${mode}.js`, { type: "text/javascript", body: await bundle(mode) })
rawInputs.forEach(([, data], i) => files.set(`/raw-${i}.bin`, { type: "application/octet-stream", body: data }))
// the swap took: only the TypeScript bundles hold the archive adapter of the TypeScript inflater
for (const mode of modes) {
  if ((files.get(`/${mode}.js`)!.body as string).includes("inflateRawSync") !== (mode !== "default")) {
    throw new Error(`bundle ${mode}: the inflater is not the one expected`)
  }
}

const server = createServer((req, res) => {
  const file = files.get(req.url ?? "")
  if (file === undefined) return void res.writeHead(404).end()
  res.writeHead(200, { "content-type": file.type }).end(file.body)
})
await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve))
const address = server.address()
const origin = `http://127.0.0.1:${typeof address === "object" && address !== null ? address.port : 0}`

// ---- CPU of the browser: user+system ticks of this process's descendants ----
const ticksPerSecond = 100
const childCpuMs = (): number => {
  let parents: Map<number, number>
  const ticks = new Map<number, number>()
  try {
    parents = new Map()
    for (const name of readdirSync("/proc")) {
      if (!/^\d+$/.test(name)) continue
      try {
        const stat = readFileSync(`/proc/${name}/stat`, "utf8")
        const fields = stat.slice(stat.lastIndexOf(")") + 2).split(" ")
        parents.set(Number(name), Number(fields[1]))
        ticks.set(Number(name), Number(fields[11]) + Number(fields[12]))
      } catch {
        // the process ended meanwhile
      }
    }
  } catch {
    return Number.NaN
  }
  let total = 0
  for (const [pid, ticksUsed] of ticks) {
    let at = parents.get(pid)
    // progress: each step moves to the parent; bounded by the process count
    for (let depth = 0; at !== undefined && at > 1 && depth < parents.size; depth++) {
      if (at === process.pid) {
        total += ticksUsed
        break
      }
      at = parents.get(at)
    }
  }
  return (total * 1000) / ticksPerSecond
}

const median = (xs: ReadonlyArray<number>) => {
  const sorted = [...xs].sort((a, b) => a - b)
  const mid = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 1 ? sorted[mid]! : (sorted[mid - 1]! + sorted[mid]!) / 2
}
const spread = (xs: ReadonlyArray<number>) => `${Math.round(Math.min(...xs))}-${Math.round(Math.max(...xs))}`
const change = (ts: number, base: number) => `${ts >= base ? "+" : ""}${((ts / base - 1) * 100).toFixed(1)}%`

interface ReadResult {
  readonly rows: number
  readonly hash: number
  readonly ms: number
  readonly maxGapMs: number
  readonly maxLongTaskMs: number
}
type Bench = Record<string, {
  readonly readRows: (bytes: Uint8Array, hash: boolean) => Promise<ReadResult>
  readonly rawInflate: (
    data: Uint8Array,
    which: "ts" | "web"
  ) => Promise<{ readonly bytes: number; readonly chunks: number; readonly ms: number }>
}>
type PageGlobals = { inflateBench: Bench; benchFiles: Record<string, Uint8Array> }

for (const name of browsers) {
  const type = playwright[name]
  if (type === undefined) throw new Error(`unknown browser ${name}`)
  const loadBefore = loadavg()
  const browser = await type.launch()
  const tab = await browser.newPage()
  await tab.goto(`${origin}/`)
  await tab.waitForFunction(modes.map((mode) => `globalThis.inflateBench?.[${JSON.stringify(mode)}] !== undefined`).join(" && "))
  await tab.evaluate(async (paths: ReadonlyArray<string>) => {
    const g = globalThis as unknown as PageGlobals
    g.benchFiles = {}
    for (const path of paths) g.benchFiles[path] = new Uint8Array(await (await fetch(path)).arrayBuffer())
  }, ["/big.xlsx", ...rawInputs.map((_, i) => `/raw-${i}.bin`)])

  const read = async (mode: string, hash: boolean) => {
    const cpuStart = childCpuMs()
    const result = await tab.evaluate(
      ([m, h]: readonly [string, boolean]) => {
        const g = globalThis as unknown as PageGlobals
        return g.inflateBench[m]!.readRows(g.benchFiles["/big.xlsx"]!, h)
      },
      [mode, hash] as const
    )
    return { ...result, cpuMs: childCpuMs() - cpuStart }
  }

  // equality, then one warm-up run per mode, untimed
  const checks: Record<string, ReadResult> = {}
  for (const mode of modes) checks[mode] = await read(mode, true)
  for (const mode of modes) await read(mode, false)
  const runs: Record<string, Array<ReadResult & { readonly cpuMs: number }>> = Object.fromEntries(modes.map((m) => [m, []]))
  // progress: k up to rounds
  for (let k = 0; k < rounds; k++) {
    // the order rotates, so no mode always runs first or after the same one
    const order = modes.map((_, j) => modes[(j + k) % modes.length]!)
    for (const mode of k % 2 === 0 ? order : order.reverse()) runs[mode]!.push(await read(mode, false))
  }

  const raw: Array<Record<string, unknown>> = []
  for (const [i, [label]] of rawInputs.entries()) {
    const times: Record<"ts" | "web", Array<number>> = { ts: [], web: [] }
    let bytes = 0
    let webChunks = 0
    // progress: k up to rawRounds
    for (let k = 0; k < rawRounds; k++) {
      const order = k % 2 === 0 ? (["web", "ts"] as const) : (["ts", "web"] as const)
      for (const which of order) {
        const r = await tab.evaluate(([path, w]: readonly [string, "ts" | "web"]) => {
          const g = globalThis as unknown as PageGlobals
          return g.inflateBench["default"]!.rawInflate(g.benchFiles[path]!, w)
        }, [`/raw-${i}.bin`, which] as const)
        bytes = r.bytes
        if (which === "web") webChunks = r.chunks
        times[which].push(r.ms)
      }
    }
    const mbps = (xs: ReadonlyArray<number>) => Math.round(bytes / 1e6 / (median(xs) / 1000))
    raw.push({
      input: label,
      outputBytes: bytes,
      tsMBps: mbps(times.ts),
      decompressionStreamMBps: mbps(times.web),
      tsTimeChange: change(median(times.ts), median(times.web)),
      decompressionStreamChunks: webChunks
    })
  }
  const version = browser.version()
  await browser.close()

  const summary = (list: ReadonlyArray<ReadResult & { readonly cpuMs: number }>) => ({
    medianMs: Math.round(median(list.map((r) => r.ms))),
    rangeMs: spread(list.map((r) => r.ms)),
    medianCpuMs: Math.round(median(list.map((r) => r.cpuMs))),
    rangeCpuMs: spread(list.map((r) => r.cpuMs)),
    medianMaxGapMs: Math.round(median(list.map((r) => r.maxGapMs))),
    worstGapMs: Math.round(Math.max(...list.map((r) => r.maxGapMs))),
    medianMaxLongTaskMs: Math.round(median(list.map((r) => r.maxLongTaskMs))),
    worstLongTaskMs: Math.round(Math.max(...list.map((r) => r.maxLongTaskMs)))
  })
  const summaries = Object.fromEntries(modes.map((mode) => [mode, summary(runs[mode]!)]))
  const base = summaries["default"]!
  const reference = checks["default"]!
  console.log(JSON.stringify({
    browser: name,
    version,
    rounds,
    setImmediate: setImmediateShim !== "",
    loadAverage: { before: loadBefore.map((x) => x.toFixed(1)), after: loadavg().map((x) => x.toFixed(1)) },
    equal: modes.every((mode) => checks[mode]!.rows === reference.rows && checks[mode]!.hash === reference.hash),
    rows: Object.fromEntries(modes.map((mode) => [mode, checks[mode]!.rows])),
    hash: Object.fromEntries(modes.map((mode) => [mode, checks[mode]!.hash.toString(16)])),
    read: summaries,
    wallChange: Object.fromEntries(modes.slice(1).map((m) => [m, change(summaries[m]!.medianMs, base.medianMs)])),
    cpuChange: Object.fromEntries(modes.slice(1).map((m) => [m, change(summaries[m]!.medianCpuMs, base.medianCpuMs)])),
    raw,
    runs: Object.fromEntries(
      Object.entries(runs).map(([mode, list]) => [mode, list.map((r) => [Math.round(r.ms), Math.round(r.cpuMs), Math.round(r.maxGapMs)])])
    )
  }))
}
server.close()

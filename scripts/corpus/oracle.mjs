#!/usr/bin/env node
// SheetJS oracle: for each manifest file, write the expected output to
// test/corpus/.cache/expected/<sha256>.json (shape documented in docs/corpus.md).
//
// SheetJS is loaded from the local clone ~/sheetjs (override: SHEETJS_DIR), never
// from node_modules. Each file is parsed in a pooled child process with a
// per-file timeout; a timeout kills the child and is recorded as such.
//
// Also probes XLSX.stream.to_json(ws, {header:1, skipHidden:true}) on files that
// have hidden rows (manifest features.hiddenRows), to confirm the known SheetJS
// hang (a hidden row does `continue` before `++R`). Probe results go to
// test/corpus/.cache/expected/<sha256>.stream.json and the run report.
//
// Usage: node scripts/corpus/oracle.mjs [--force] [--jobs N] [--timeout SEC]
//          [--probe-timeout SEC] [--no-probe] [--source sheetjs|generated]
//          [--format xls,xlsb,...] [--max-size BYTES] [--limit N]
// Writes a run report to test/corpus/.cache/oracle-report.json.

import { fork } from "node:child_process"
import { Worker } from "node:worker_threads"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { createRequire } from "node:module"
import { fileURLToPath } from "node:url"
import { HIDDEN_ROWS_DIR, CACHE_DIR, EXPECTED_DIR, GENERATED_DIR, MANIFEST_PATH, SHEETJS_CACHE, SHEETJS_DIR, SHEETJS_ENTRY } from "./lib.mjs"

export const SCHEMA = 1
// Options passed to XLSX.read. Recorded in every expected file.
export const READ_OPTIONS = {
  type: "buffer",
  dense: true, // memory scales with present cells, also for huge declared ranges
  cellDates: false, // dates stay serial numbers (v) with their format (z); our reader flags dates itself
  cellNF: true, // keep number format code in z
  cellFormula: true, // keep formulas in f
  cellText: true, // keep SheetJS formatted text in w (for later SSF comparisons)
  cellStyles: true, // needed for XLSX to record hidden rows in !rows (without it the xlsx reader drops the flag)
  sheetStubs: false,
  WTF: false,
}
// sheet_to_json options for `rows` (only computed when the !ref area is small enough).
export const JSON_OPTIONS = { header: 1, raw: true, defval: null, blankrows: true }
export const ROWS_MAX_AREA = 2_000_000
// Caps per expected file: above these the file holds status "tooBig" and counts only.
export const MAX_CELLS = 5_000_000
export const MAX_BYTES = 128 * 1024 * 1024
// A typical SheetJS-based reader: XLSX.read(buffer) then stream.to_json with these options.
export const PROBE_OPTIONS = { raw: true, skipHidden: true, UTC: true, header: "A" }
// Probe variants:
//   default    = XLSX.read(buf) as the old reader did, then stream.to_json(PROBE_OPTIONS)
//   cellStyles = same with cellStyles:true (needed for xlsx to record hidden rows)
//   control    = cellStyles:true but skipHidden:false: if this finishes and the others time out,
//                the timeout is the skipHidden hang, not just a slow range scan.
export const PROBE_VARIANTS = {
  default: { read: { type: "buffer" }, stream: PROBE_OPTIONS },
  cellStyles: { read: { type: "buffer", cellStyles: true }, stream: PROBE_OPTIONS },
  control: { read: { type: "buffer", cellStyles: true }, stream: { ...PROBE_OPTIONS, skipHidden: false } },
}

const SRC_ROOT = { sheetjs: path.join(SHEETJS_CACHE, "test_files"), generated: GENERATED_DIR, "hidden-rows": HIDDEN_ROWS_DIR }

// ------------------------------------------------------------------ child

// Watchdog thread for each child. It runs even while the child's main thread is
// stuck in a SheetJS loop, and SIGKILLs the child when (a) the parent goes away
// (ppid changes, e.g. the parent was SIGKILLed alone) or (b) the hard deadline of
// the current job passes (job timeout + grace; the parent normally kills first).
const WATCHDOG = `
const { parentPort, workerData } = require("node:worker_threads")
const ppid = workerData.ppid
let deadline = null
parentPort.on("message", (m) => { deadline = m.deadline })
setInterval(() => {
  if (process.ppid !== ppid || (deadline !== null && Date.now() > deadline)) process.kill(process.pid, "SIGKILL")
}, 250)
`
const HARD_GRACE_MS = 10_000

if (process.argv[2] === "--child") {
  const watchdog = new Worker(WATCHDOG, { eval: true, workerData: { ppid: process.ppid } })
  watchdog.unref()
  // If the IPC channel closes while we're idle, leave at once.
  process.on("disconnect", () => process.exit(0))
  const require = createRequire(import.meta.url)
  const XLSX = require(SHEETJS_ENTRY)
  const meta = { oracle: "sheetjs", sheetjsVersion: XLSX.version, sheetjsCommit: sheetjsCommit() }
  process.on("message", async (job) => {
    watchdog.postMessage({ deadline: Date.now() + job.timeout + HARD_GRACE_MS })
    const t0 = performance.now()
    let out
    try {
      out = job.kind === "probe" ? await probe(XLSX, job) : parse(XLSX, job)
    } catch (e) {
      out = { status: "error", error: String(e?.message ?? e).slice(0, 2000) }
    }
    const ms = Math.round(performance.now() - t0)
    const head = { schema: SCHEMA, ...meta, sha256: job.sha256, path: job.path, source: job.source, format: job.format, kind: job.kind, ms }
    if (job.variant) head.variant = job.variant
    try {
      writeJson(job.out, { ...head, ...out }, MAX_BYTES)
    } catch (e) {
      out = e?.code === "TOO_BIG"
        ? tooBig(out, `expected JSON over ${MAX_BYTES} bytes`)
        : { status: "error", error: `could not serialize result: ${String(e?.message ?? e).slice(0, 500)}` }
      fs.writeFileSync(job.out, JSON.stringify({ ...head, ...out }) + "\n")
    }
    watchdog.postMessage({ deadline: null })
    process.send({ id: job.id, status: out.status, ms, error: out.error })
  })
  process.send({ ready: true })
}

function sheetjsCommit() {
  try {
    const head = fs.readFileSync(path.join(SHEETJS_DIR, ".git", "HEAD"), "utf8").trim()
    if (!head.startsWith("ref:")) return head.slice(0, 12)
    const ref = head.slice(5)
    const p = path.join(SHEETJS_DIR, ".git", ref)
    if (fs.existsSync(p)) return fs.readFileSync(p, "utf8").trim().slice(0, 12)
    const packed = fs.readFileSync(path.join(SHEETJS_DIR, ".git", "packed-refs"), "utf8")
    const m = packed.split("\n").find((l) => l.endsWith(" " + ref))
    return m ? m.slice(0, 12) : null
  } catch { return null }
}

// Streaming JSON writer: the result for big sheets can exceed V8's max string
// length, so big arrays are written element by element (same output as
// JSON.stringify(value, replacer)).
// Throws an error with code "TOO_BIG" once more than maxBytes (UTF-16 units, close
// enough) would be written; the caller then replaces the file.
function writeJson(file, value, maxBytes = Infinity) {
  const fd = fs.openSync(file, "w")
  let pending = []
  let size = 0, total = 0
  const emit = (s) => {
    pending.push(s)
    size += s.length
    total += s.length
    if (total > maxBytes) throw Object.assign(new Error("too big"), { code: "TOO_BIG" })
    if (size > 1 << 22) { fs.writeSync(fd, pending.join("")); pending = []; size = 0 }
  }
  // depth 0 doc, 1 sheets[], 2 sheet{}, 3 cells[]/rows[] (walked when long), 4 cell/row (stringified)
  const walk = (v, depth) => {
    if (Array.isArray(v) && (depth < 3 || v.length > 1000)) {
      emit("[")
      for (let i = 0; i < v.length; i++) { if (i) emit(","); walk(v[i] === undefined ? null : v[i], depth + 1) }
      emit("]")
    } else if (v && typeof v === "object" && !Array.isArray(v) && !(v instanceof Date) && depth < 3) {
      emit("{")
      let first = true
      for (const [k, x] of Object.entries(v)) {
        if (x === undefined) continue
        if (!first) emit(",")
        first = false
        emit(JSON.stringify(k) + ":")
        walk(x, depth + 1)
      }
      emit("}")
    } else emit(JSON.stringify(v, replacer) ?? "null")
  }
  try {
    walk(value, 0)
    emit("\n")
    fs.writeSync(fd, pending.join(""))
  } finally { fs.closeSync(fd) }
}

// JSON cannot carry NaN/Infinity/undefined/Date: encode explicitly.
function replacer(_k, v) {
  if (typeof v === "number" && !Number.isFinite(v)) return { $num: String(v) }
  if (v instanceof Date) return { $date: Number.isNaN(v.getTime()) ? "Invalid" : v.toISOString() }
  return v
}

const VIS = ["visible", "hidden", "veryHidden"]

function parse(XLSX, job) {
  const buf = fs.readFileSync(job.file)
  const wb = XLSX.read(buf, { ...READ_OPTIONS }) // SheetJS mutates its options: always pass a copy
  const wbSheets = wb.Workbook?.Sheets ?? []
  const workbook = {
    bookType: wb.bookType ?? null,
    date1904: !!wb.Workbook?.WBProps?.date1904,
    sheets: wb.SheetNames.map((name, i) => ({ name, visibility: VIS[wbSheets[i]?.Hidden ?? 0] ?? String(wbSheets[i]?.Hidden) })),
  }
  const options = { read: READ_OPTIONS, json: JSON_OPTIONS, rowsMaxArea: ROWS_MAX_AREA, maxCells: MAX_CELLS, maxBytes: MAX_BYTES }
  const counts = wb.SheetNames.map((name) => sheetCounts(wb.Sheets[name], name))
  const total = counts.reduce((a, c) => a + c.cellCount, 0)
  if (total > MAX_CELLS) return tooBig({ options, workbook, sheets: counts }, `${total} cells > ${MAX_CELLS}`)
  const sheets = wb.SheetNames.map((name) => sheetOut(XLSX, wb.Sheets[name], name))
  return { status: "ok", options, workbook, sheets }
}

// Counts-only view of a sheet (for "tooBig" results).
function sheetCounts(ws, name) {
  if (!ws) return { name, missing: true, cellCount: 0 }
  let cellCount = 0
  const data = ws["!data"]
  if (data) { for (const row of data) if (row) for (const c of row) if (c && c.t != null) cellCount++ }
  else for (const k of Object.keys(ws)) if (k[0] !== "!" && ws[k]?.t != null) cellCount++
  const hiddenRows = (ws["!rows"] ?? []).reduce((a, r) => a + (r?.hidden ? 1 : 0), 0)
  return { name, ref: ws["!ref"] ?? null, cellCount, hiddenRowCount: hiddenRows }
}

function tooBig(out, reason) {
  return {
    status: "tooBig", reason, options: out.options, workbook: out.workbook,
    sheets: (out.sheets ?? []).map((s) => ({ name: s.name, ref: s.ref ?? null, cellCount: s.cellCount ?? 0, hiddenRowCount: s.hiddenRowCount ?? s.hiddenRows?.length ?? 0 })),
  }
}

function sheetOut(XLSX, ws, name) {
  if (!ws) return { name, missing: true }
  const ref = ws["!ref"] ?? null
  const o = { name, ref }
  if (ws["!fullref"] && ws["!fullref"] !== ref) o.fullref = ws["!fullref"]
  if (ws["!type"]) o.type = ws["!type"]
  const hiddenRows = [], hiddenCols = []
  ;(ws["!rows"] ?? []).forEach((r, i) => { if (r?.hidden) hiddenRows.push(i) })
  ;(ws["!cols"] ?? []).forEach((c, i) => { if (c?.hidden) hiddenCols.push(i) })
  o.hiddenRows = hiddenRows
  o.hiddenCols = hiddenCols
  o.merges = (ws["!merges"] ?? []).map((m) => XLSX.utils.encode_range(m))
  // Sparse cell list: every cell object SheetJS produced, row-major.
  const cells = []
  const push = (r, c, cell) => {
    if (!cell || cell.t == null) return
    const x = { r, c, t: cell.t }
    if (cell.v !== undefined) x.v = cell.v
    if (cell.z !== undefined && cell.z !== "General") x.z = cell.z
    if (cell.w !== undefined) x.w = cell.w
    if (cell.f !== undefined) x.f = cell.f
    if (cell.F !== undefined) x.F = cell.F
    cells.push(x)
  }
  const data = ws["!data"]
  if (data) {
    for (let r = 0; r < data.length; r++) {
      const row = data[r]
      if (!row) continue
      for (let c = 0; c < row.length; c++) push(r, c, row[c])
    }
  } else {
    for (const k of Object.keys(ws)) {
      if (k[0] === "!") continue
      const a = XLSX.utils.decode_cell(k)
      push(a.r, a.c, ws[k])
    }
    cells.sort((a, b) => a.r - b.r || a.c - b.c)
  }
  o.cellCount = cells.length
  o.cells = cells
  if (ref) {
    const rg = XLSX.utils.decode_range(ref)
    const area = (rg.e.r - rg.s.r + 1) * (rg.e.c - rg.s.c + 1)
    if (area <= ROWS_MAX_AREA) o.rows = XLSX.utils.sheet_to_json(ws, { ...JSON_OPTIONS })
    else { o.rows = null; o.rowsOmitted = `ref area ${area} > ${ROWS_MAX_AREA}` }
  } else o.rows = []
  return o
}

async function probe(XLSX, job) {
  const buf = fs.readFileSync(job.file)
  const v = PROBE_VARIANTS[job.variant]
  const wb = XLSX.read(buf, { ...v.read }) // SheetJS mutates its options: always pass a copy
  const sheets = []
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name]
    // If stream.to_json's row loop never advances, _read never returns and the
    // parent kills us on timeout.
    let rows = 0
    for await (const _row of XLSX.stream.to_json(ws, { ...v.stream })) rows++
    sheets.push({ name, rows })
  }
  return { status: "ok", variant: job.variant, options: v, sheets }
}

// ------------------------------------------------------------------ parent

// Live children, SIGKILLed when the parent exits or is interrupted.
const live = new Set()
const killAll = () => { for (const c of live) { try { c.kill("SIGKILL") } catch {} } live.clear() }
process.on("exit", killAll)
for (const [sig, code] of [["SIGINT", 130], ["SIGTERM", 143], ["SIGHUP", 129]]) {
  process.on(sig, () => { killAll(); process.exit(code) })
}
const MAX_STARTUP_DEATHS = 3

if (process.argv[2] !== "--child") await main()

async function main() {
  const args = parseArgs(process.argv.slice(2))
  if (!fs.existsSync(SHEETJS_ENTRY)) {
    console.error(`SheetJS not found at ${SHEETJS_DIR} (set SHEETJS_DIR); skipping oracle`)
    process.exit(0)
  }
  const manifest = JSON.parse(fs.readFileSync(MANIFEST_PATH, "utf8"))
  fs.mkdirSync(EXPECTED_DIR, { recursive: true })
  const formats = args.format ? new Set(args.format.split(",")) : null
  // One job per unique sha256 (many corpus files are byte-identical copies).
  const seen = new Set()
  const jobs = []
  for (const e of manifest.files) {
    if (args.source && e.source !== args.source) continue
    if (formats && !formats.has(e.format)) continue
    if (args.maxSize && e.size > args.maxSize) continue
    if (e.format === "ooxml-encrypted") continue // SheetJS can't decrypt; nothing to compare
    const file = path.join(SRC_ROOT[e.source], e.path)
    if (!fs.existsSync(file)) continue
    const base = { sha256: e.sha256, path: e.path, source: e.source, format: e.format, file, size: e.size }
    if (!seen.has(e.sha256)) {
      seen.add(e.sha256)
      jobs.push({ ...base, kind: "parse", out: path.join(EXPECTED_DIR, `${e.sha256}.json`), timeout: args.timeout })
    }
    if (args.probe && (e.features?.hiddenRows || e.source !== "sheetjs") && !seen.has("p:" + e.sha256)) {
      seen.add("p:" + e.sha256)
      for (const variant of Object.keys(PROBE_VARIANTS)) {
        jobs.push({ ...base, kind: "probe", variant, out: path.join(EXPECTED_DIR, `${e.sha256}.stream.${variant}.json`), timeout: args.probeTimeout })
      }
    }
  }
  const todo = (args.limit ? jobs.slice(0, args.limit) : jobs).filter((j) => args.force || !upToDate(j.out))
  // Big files first so the long tail overlaps with small ones.
  todo.sort((a, b) => b.size - a.size)
  console.log(`oracle: ${jobs.length} jobs, ${todo.length} to run, ${args.jobs} workers, timeout ${args.timeout / 1000}s (probe ${args.probeTimeout / 1000}s)`)
  const t0 = Date.now()
  let results
  try {
    results = await runPool(todo, args.jobs, args.readyTimeout)
  } catch (e) {
    console.error(String(e.message ?? e))
    process.exit(2)
  }
  // Report covers all jobs (fresh + cached).
  const report = { schema: SCHEMA, generatedAt: new Date().toISOString(), wallMs: Date.now() - t0, options: { read: READ_OPTIONS, json: JSON_OPTIONS, probe: PROBE_VARIANTS, timeoutMs: args.timeout, probeTimeoutMs: args.probeTimeout }, counts: {}, notOk: [], probes: [] }
  for (const j of jobs) {
    let r = results.get(j)
    if (!r) {
      try { const d = JSON.parse(fs.readFileSync(j.out, "utf8")); r = { status: d.status, ms: d.ms, error: d.error } } catch { continue }
    }
    const key = j.kind === "probe" ? `probe-${j.variant}:${j.format}` : `parse:${j.format}`
    const c = (report.counts[key] ??= { ok: 0, tooBig: 0, error: 0, timeout: 0, crash: 0 })
    c[r.status] = (c[r.status] ?? 0) + 1
    if (j.kind === "probe") report.probes.push({ path: j.path, source: j.source, format: j.format, sha256: j.sha256, variant: j.variant, status: r.status, ms: r.ms, error: r.error })
    else if (r.status !== "ok" && r.status !== "tooBig") report.notOk.push({ path: j.path, source: j.source, format: j.format, sha256: j.sha256, status: r.status, ms: r.ms, error: r.error })
  }
  // Per-file verdict on the stream.to_json(skipHidden) hang.
  const byFile = new Map()
  for (const p of report.probes) {
    const k = `${p.source}/${p.path}`
    if (!byFile.has(k)) byFile.set(k, { file: k, format: p.format, sha256: p.sha256 })
    byFile.get(k)[p.variant] = p.status
  }
  report.streamVerdicts = [...byFile.values()].map((f) => {
    const hung = f.default === "timeout" || f.cellStyles === "timeout"
    f.verdict = !hung ? (f.default === "error" ? "error" : "ok")
      : f.control === "ok" ? "hang" // skipHidden:false finishes, skipHidden:true doesn't
      : "slow" // control also times out: range scan cost (huge !ref), not the hidden-row loop
    return f
  })
  const verdicts = {}
  for (const f of report.streamVerdicts) verdicts[f.verdict] = (verdicts[f.verdict] ?? 0) + 1
  report.streamVerdictCounts = verdicts
  fs.writeFileSync(path.join(CACHE_DIR, "oracle-report.json"), JSON.stringify(report, null, 2) + "\n")
  console.log(`done in ${((Date.now() - t0) / 1000).toFixed(1)}s`)
  console.table(report.counts)
  if (report.notOk.length) {
    console.log("not ok (parse):")
    for (const n of report.notOk) console.log(`  ${n.status.padEnd(7)} ${n.format.padEnd(6)} ${n.source}/${n.path}${n.error ? "  — " + n.error.slice(0, 120) : ""}`)
  }
  const pt = report.probes.filter((p) => p.status !== "ok")
  console.log(`stream.to_json probe: ${report.probes.length} probes on ${report.streamVerdicts.length} files, ${pt.length} not ok; verdicts`, verdicts)
  if (args.verbose) for (const p of pt) console.log(`  ${p.status.padEnd(7)} ${p.variant.padEnd(10)} ${p.format.padEnd(6)} ${p.source}/${p.path}${p.error && p.status !== "timeout" ? "  — " + p.error.slice(0, 100) : ""}`)
  for (const f of report.streamVerdicts) if (f.verdict !== "ok") console.log(`  ${f.verdict.padEnd(5)} ${f.format.padEnd(5)} default=${f.default} cellStyles=${f.cellStyles} control=${f.control}  ${f.file}`)
}

function upToDate(out) {
  try { return JSON.parse(fs.readFileSync(out, "utf8")).schema === SCHEMA } catch { return false }
}

function parseArgs(argv) {
  const a = { verbose: false, force: false, jobs: Math.max(1, Math.min(10, Math.floor(os.availableParallelism() / 2))), timeout: 60_000, probeTimeout: 15_000, readyTimeout: 30_000, probe: true, source: null, format: null, maxSize: 0, limit: 0 }
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i], v = argv[i + 1]
    if (k === "--force") a.force = true
    else if (k === "--no-probe") a.probe = false
    else if (k === "--verbose") a.verbose = true
    else if (k === "--jobs") { a.jobs = +v; i++ }
    else if (k === "--timeout") { a.timeout = +v * 1000; i++ }
    else if (k === "--probe-timeout") { a.probeTimeout = +v * 1000; i++ }
    else if (k === "--ready-timeout") { a.readyTimeout = +v * 1000; i++ }
    else if (k === "--source") { a.source = v; i++ }
    else if (k === "--format") { a.format = v; i++ }
    else if (k === "--max-size") { a.maxSize = +v; i++ }
    else if (k === "--limit") { a.limit = +v; i++ }
    else throw new Error(`unknown arg ${k}`)
  }
  return a
}

function runPool(jobs, n, readyTimeout) {
  const results = new Map()
  if (!jobs.length) return Promise.resolve(results)
  let next = 0, done = 0, active = 0, startupDeaths = 0, aborted = false
  const self = fileURLToPath(import.meta.url)
  return new Promise((resolve, reject) => {
    const progress = () => { if (done % 200 === 0 || done === jobs.length) console.log(`  ${done}/${jobs.length}`) }
    const finish = (job, r) => {
      results.set(job, r)
      if (r.status !== "ok" && r.status !== "error") {
        // the child couldn't write the file: record it ourselves
        const doc = { schema: SCHEMA, oracle: "sheetjs", sha256: job.sha256, path: job.path, source: job.source, format: job.format, kind: job.kind, variant: job.variant, status: r.status, ms: r.ms, error: r.error }
        fs.writeFileSync(job.out, JSON.stringify(doc) + "\n")
      }
      done++
      progress()
    }
    const abort = (why) => {
      if (aborted) return
      aborted = true
      killAll()
      reject(new Error(why))
    }
    const spawnWorker = () => {
      if (aborted) return
      if (next >= jobs.length) { if (active === 0) resolve(results); return }
      active++
      const child = fork(self, ["--child"], { execArgv: ["--max-old-space-size=4096", "--stack-size=4000"], stdio: ["ignore", "ignore", "pipe", "ipc"] })
      live.add(child)
      let stderr = ""
      child.stderr.on("data", (d) => { stderr = (stderr + d).slice(-2000) })
      let cur = null, timer = null, started = 0, ready = false
      // A child that never says ready (stuck or slow while loading xlsx.js) is killed and counts as a startup death.
      const readyTimer = setTimeout(() => { stderr += `\n(no ready within ${readyTimeout / 1000}s)`; child.kill("SIGKILL") }, readyTimeout)
      const give = () => {
        if (aborted || next >= jobs.length) { child.kill(); return }
        cur = jobs[next++]
        cur.id = next
        started = Date.now()
        timer = setTimeout(() => {
          const j = cur; cur = null
          child.kill("SIGKILL")
          finish(j, { status: "timeout", ms: Date.now() - started, error: `no result within ${j.timeout / 1000}s` })
        }, cur.timeout)
        child.send(cur)
      }
      child.on("message", (m) => {
        if (m.ready) { ready = true; startupDeaths = 0; clearTimeout(readyTimer); return give() }
        if (!cur || m.id !== cur.id) return
        clearTimeout(timer)
        const j = cur; cur = null
        finish(j, m)
        give()
      })
      child.on("exit", (code, sig) => {
        clearTimeout(timer)
        clearTimeout(readyTimer)
        live.delete(child)
        active--
        if (!ready) {
          // Died (or was killed) before it could take a job: don't respawn forever.
          if (++startupDeaths >= MAX_STARTUP_DEATHS) return abort(`oracle: ${startupDeaths} consecutive children died at startup (last: exit ${code ?? sig}): ${stderr.trim().slice(-800)}`)
        }
        if (cur) { finish(cur, { status: "crash", ms: Date.now() - started, error: `child exited ${code ?? sig}: ${stderr.slice(-500)}` }); cur = null }
        spawnWorker()
      })
    }
    for (let i = 0; i < Math.min(n, jobs.length); i++) spawnWorker()
  })
}

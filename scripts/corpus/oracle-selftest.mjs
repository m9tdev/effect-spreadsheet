#!/usr/bin/env node
// Repro/test for oracle.mjs process hygiene. Each case runs the real oracle on the
// committed hidden-rows fixtures in a temporary cache (CORPUS_CACHE_DIR) and checks:
//   A. parent SIGKILLed alone while children hang in SheetJS: the children's
//      watchdog threads see the ppid change and kill them (no orphans).
//   B. parent SIGTERMed: the parent SIGKILLs its live children before exiting.
//   C. a throwing xlsx.js: children die at startup; the oracle aborts after 3
//      consecutive startup deaths (exit 2) instead of respawning forever.
//   D. an xlsx.js that never finishes loading: the ready timer kills each child;
//      same abort after 3.
//   E. tooBig results (tiny ORACLE_MAX_CELLS) keep reason, options, workbook,
//      per-sheet counts and the SheetJS version in the expected file.
//   F. ORACLE_MAX_BYTES counts UTF-8 bytes, not UTF-16 string length: with a cap
//      between the two sizes of generated/features.xlsx's expected JSON (it holds
//      non-ASCII text, so bytes > chars), the result must be tooBig.
//   G. the oracle output doesn't depend on the time zone. Two sub-checks on
//      generated/dates.xlsx, both compared with the committed rows:
//      G1 runs the oracle under TZ=Pacific/Auckland. Its children run with TZ=UTC, so this
//         only fails when both safeguards (sheet_to_json UTC:true and the children's
//         TZ=UTC) are removed.
//      G2 runs just the JSON step (READ_OPTIONS + JSON_OPTIONS from oracle-options.mjs)
//         in a child with TZ=Pacific/Auckland and no TZ=UTC override. This fails when
//         UTC:true alone is dropped.
//   H. a timeout is not a cached result: after a run where every parse times out
//      (--timeout 0.001), a run without --force parses those files again.
// Needs SheetJS (SHEETJS_DIR, default ~/sheetjs) for A, B and E: they are skipped, not
// failed, without it. Linux `ps` for the child lookup. Stale temp dirs from a SIGKILLed
// run are removed at the next start.
// Usage: node scripts/corpus/oracle-selftest.mjs

import { execFileSync, spawn } from "node:child_process"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const HERE = path.dirname(fileURLToPath(import.meta.url))
const { parseFlags, SHEETJS_ENTRY } = await import(path.join(HERE, "lib.mjs"))
parseFlags(process.argv.slice(2), []) // no arguments
const ORACLE = path.join(HERE, "oracle.mjs")
const MANIFEST = path.join(HERE, "manifest.mjs")
// Temp dirs carry our pid; remove ones left by a selftest that was SIGKILLed (pid gone).
for (const d of fs.readdirSync(os.tmpdir())) {
  if (!d.startsWith("oracle-selftest-")) continue
  const m = /^oracle-selftest-(\d+)-/.exec(d)
  if (m && Number(m[1]) === process.pid) continue
  let gone = false
  if (m) { try { process.kill(Number(m[1]), 0) } catch (e) { gone = e.code === "ESRCH" } }
  else { try { gone = Date.now() - fs.statSync(path.join(os.tmpdir(), d)).mtimeMs > 3_600_000 } catch {} } // older name format: by age
  if (gone) fs.rmSync(path.join(os.tmpdir(), d), { recursive: true, force: true })
}
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), `oracle-selftest-${process.pid}-`))
for (const sig of ["SIGINT", "SIGTERM", "SIGHUP"]) process.on(sig, () => { fs.rmSync(tmp, { recursive: true, force: true }); process.exit(1) })
const haveSheetJS = fs.existsSync(SHEETJS_ENTRY)
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const alive = (pid) => { try { process.kill(pid, 0); return true } catch { return false } }
const childrenOf = (pid) => {
  try { return execFileSync("ps", ["-o", "pid=", "--ppid", String(pid)], { encoding: "utf8" }).split(/\s+/).filter(Boolean).map(Number) } catch { return [] }
}

let failed = 0
const skip = (name, why) => console.log(`skip ${name}  — ${why}`)
const check = (name, ok, detail = "") => { console.log(`${ok ? "ok  " : "FAIL"} ${name}${detail ? "  — " + detail : ""}`); if (!ok) failed++ }

function startOracle(env, args) {
  const p = spawn(process.execPath, [ORACLE, "--source", "hidden-rows", "--force", ...args], { env: { ...process.env, CORPUS_CACHE_DIR: tmp, ...env }, stdio: ["ignore", "pipe", "pipe"] })
  let out = ""
  p.stdout.on("data", (d) => { out += d })
  p.stderr.on("data", (d) => { out += d })
  const exited = new Promise((r) => p.on("exit", (code, sig) => r({ code, sig, out })))
  return { p, exited, output: () => out }
}

// Exit code the parent's own handler uses (128 + signal number); SIGKILL can't be handled.
const HANDLED_EXIT = { SIGTERM: 143, SIGINT: 130, SIGHUP: 129 }

async function killCase(name, signal) {
  // Probes on hidden-rows/hidden-row.* hang in SheetJS; give them a long timeout so they're still hanging when we kill.
  const o = startOracle({}, ["--jobs", "3", "--probe-timeout", "120", "--timeout", "120"])
  let kids = []
  for (let i = 0; i < 40 && kids.length < 3; i++) { await sleep(250); kids = childrenOf(o.p.pid) }
  await sleep(2000) // let them get into the hanging probes
  kids = childrenOf(o.p.pid)
  check(`${name}: oracle has children`, kids.length > 0, `pids ${kids.join(",")}`)
  o.p.kill(signal)
  const r = await o.exited
  // The watchdogs would clean up even without the parent's handler, so check the handler ran.
  if (HANDLED_EXIT[signal]) check(`${name}: parent's ${signal} handler ran (exit ${HANDLED_EXIT[signal]})`, r.code === HANDLED_EXIT[signal], `exit ${r.code ?? r.sig}`)
  let left = kids
  for (let i = 0; i < 20 && left.length; i++) { await sleep(250); left = left.filter(alive) }
  check(`${name}: no children left 5 s after ${signal} of the parent`, left.length === 0, left.length ? `still alive: ${left.join(",")}` : "")
  for (const pid of left) { try { process.kill(pid, "SIGKILL") } catch {} }
}

async function tooBigCase(name) {
  const o = startOracle({ ORACLE_MAX_CELLS: "5" }, ["--no-probe", "--format", "xlsx", "--jobs", "2"])
  const r = await o.exited
  check(`${name}: oracle exits 0`, r.code === 0, `exit ${r.code ?? r.sig}`)
  const dir = path.join(tmp, "expected")
  // --force rewrote every xlsx parse result in this run; earlier cases may have left other formats.
  const docs = fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => !f.includes(".stream.")).map((f) => JSON.parse(fs.readFileSync(path.join(dir, f), "utf8"))).filter((d) => d.format === "xlsx") : []
  const big = docs.filter((d) => d.status === "tooBig")
  check(`${name}: xlsx results are tooBig`, big.length > 0 && big.length === docs.length, `${big.length}/${docs.length}`)
  const complete = big.filter((d) => /cells > 5$/.test(d.reason ?? "") && d.options?.maxCells === 5 && d.workbook?.sheets?.length && d.sheets?.length && d.sheets.every((s) => typeof s.cellCount === "number" && "ref" in s && typeof s.hiddenRowCount === "number") && d.sheetjsVersion)
  check(`${name}: every tooBig doc keeps reason, options, workbook, counts, sheetjsVersion`, big.length > 0 && complete.length === big.length, big[0] ? Object.keys(big[0]).join(",") : "")
}

async function maxBytesCase(name) {
  const docFor = () => {
    const dir = path.join(tmp, "expected")
    for (const f of fs.readdirSync(dir)) {
      if (f.includes(".stream.")) continue
      const p = path.join(dir, f), d = JSON.parse(fs.readFileSync(p, "utf8"))
      if (d.source === "generated" && d.path === "features.xlsx") return { p, d }
    }
    return null
  }
  const args = ["--source", "generated", "--format", "xlsx", "--no-probe", "--jobs", "2"]
  let r = await startOracle({}, args).exited
  const first = docFor()
  check(`${name}: features.xlsx parsed without a cap`, r.code === 0 && first?.d.status === "ok", `exit ${r.code ?? r.sig}, status ${first?.d.status}`)
  if (!first) return
  const text = fs.readFileSync(first.p, "utf8")
  const bytes = Buffer.byteLength(text, "utf8"), chars = text.length
  check(`${name}: expected JSON has more bytes than chars`, bytes - chars >= 40, `${bytes} bytes, ${chars} chars`)
  const cap = Math.floor((bytes + chars) / 2) // over the char count, under the byte count
  r = await startOracle({ ORACLE_MAX_BYTES: String(cap) }, args).exited
  const second = docFor()
  check(`${name}: cap ${cap} (between ${chars} chars and ${bytes} bytes) gives tooBig`, second?.d.status === "tooBig" && /over \d+ bytes/.test(second.d.reason ?? ""), `status ${second?.d.status}, reason ${second?.d.reason}`)
}

async function timeZoneCase(name) {
  const r = await startOracle({ TZ: "Pacific/Auckland" }, ["--source", "generated", "--format", "xlsx", "--no-probe", "--jobs", "2"]).exited
  const dir = path.join(tmp, "expected")
  const doc = fs.readdirSync(dir).filter((f) => !f.includes(".stream.")).map((f) => JSON.parse(fs.readFileSync(path.join(dir, f), "utf8"))).find((d) => d.source === "generated" && d.path === "dates.xlsx")
  const committed = JSON.parse(fs.readFileSync(path.join(HERE, "..", "..", "test", "corpus", "expected", "generated", "dates.xlsx.json"), "utf8"))
  const want = JSON.stringify(committed.sheets.map((x) => x.rows))
  check(`${name}1 time zone: oracle under TZ=Pacific/Auckland gives the committed dates.xlsx rows`, r.code === 0 && !!doc && JSON.stringify(doc.sheets.map((x) => x.rows)) === want, doc ? `first value ${JSON.stringify(doc.sheets[0].rows?.[0]?.[0])}` : "no result")
  // G2: the JSON step alone, in a child that really runs in Pacific/Auckland (no TZ=UTC override).
  const code = `
    import fs from "node:fs"
    import { createRequire } from "node:module"
    const { READ_OPTIONS, JSON_OPTIONS } = await import(${JSON.stringify(new URL("./oracle-options.mjs", import.meta.url).href)})
    const XLSX = createRequire(import.meta.url)(${JSON.stringify(SHEETJS_ENTRY)})
    const wb = XLSX.read(fs.readFileSync(${JSON.stringify(path.join(HERE, "..", "..", "test", "corpus", "generated", "dates.xlsx"))}), { ...READ_OPTIONS })
    process.stdout.write(JSON.stringify(wb.SheetNames.map((n) => XLSX.utils.sheet_to_json(wb.Sheets[n], { ...JSON_OPTIONS }))))
  `
  const env = { ...process.env, TZ: "Pacific/Auckland" }
  const out = execFileSync(process.execPath, ["--input-type=module", "-e", code], { env, encoding: "utf8" })
  const got = JSON.parse(out)
  check(`${name}2 time zone: sheet_to_json with JSON_OPTIONS in a TZ=Pacific/Auckland child gives the committed rows`, JSON.stringify(got) === want, `first value ${JSON.stringify(got[0]?.[0]?.[0])}`)
}

async function retryTimeoutCase(name) {
  const args = ["--no-probe", "--jobs", "2"]
  const dir = path.join(tmp, "expected")
  const statuses = () => fs.readdirSync(dir).filter((f) => !f.includes(".stream.")).map((f) => JSON.parse(fs.readFileSync(path.join(dir, f), "utf8"))).filter((d) => d.source === "hidden-rows").map((d) => d.status)
  const r = await startOracle({}, [...args, "--timeout", "0.001"]).exited
  const first = statuses()
  check(`${name}: a run with --timeout 0.001 times out`, r.code === 0 && first.length > 0 && first.every((s) => s === "timeout"), first.join(","))
  // startOracle passes --force; run the oracle directly without it
  const p = spawn(process.execPath, [ORACLE, "--source", "hidden-rows", ...args], { env: { ...process.env, CORPUS_CACHE_DIR: tmp }, stdio: ["ignore", "pipe", "pipe"] })
  let out = ""
  p.stdout.on("data", (d) => { out += d }); p.stderr.on("data", (d) => { out += d })
  const code = await new Promise((res) => p.on("exit", (c) => res(c)))
  const m = /oracle: (\d+) jobs, (\d+) to run/.exec(out)
  check(`${name}: the next run without --force retries them`, code === 0 && !!m && +m[2] > 0 && +m[2] === +m[1], m ? `${m[2]} of ${m[1]} jobs to run` : out.slice(0, 120))
  const second = statuses()
  check(`${name}: and they are ok now`, second.length > 0 && second.every((s) => s === "ok"), second.join(","))
}

async function startupCase(name, body, args, expect) {
  const dir = path.join(tmp, name.replace(/\W+/g, "_"))
  fs.mkdirSync(dir, { recursive: true })
  fs.writeFileSync(path.join(dir, "xlsx.js"), body)
  const t0 = Date.now()
  const o = startOracle({ SHEETJS_DIR: dir }, ["--no-probe", "--jobs", "2", ...args])
  const killer = setTimeout(() => o.p.kill("SIGKILL"), 60_000)
  const r = await o.exited
  clearTimeout(killer)
  const ms = Date.now() - t0
  check(`${name}: aborts with exit 2`, r.code === 2, `exit ${r.code ?? r.sig} after ${ms} ms`)
  check(`${name}: reports 3 startup deaths`, /3 consecutive children died at startup/.test(r.out), r.out.trim().split("\n").pop().slice(0, 160))
  check(`${name}: error mentions ${expect}`, r.out.includes(expect))
}

try {
  execFileSync(process.execPath, [MANIFEST], { env: { ...process.env, CORPUS_CACHE_DIR: tmp }, stdio: "ignore" })
  if (haveSheetJS) {
    await killCase("A parent SIGKILL", "SIGKILL")
    await killCase("B parent SIGTERM", "SIGTERM")
  } else {
    skip("A parent SIGKILL", `needs SheetJS at ${SHEETJS_ENTRY} (set SHEETJS_DIR)`)
    skip("B parent SIGTERM", `needs SheetJS at ${SHEETJS_ENTRY} (set SHEETJS_DIR)`)
  }
  await startupCase("C throwing xlsx.js", 'throw new Error("boom from fake xlsx.js")\n', [], "boom from fake xlsx.js")
  await startupCase("D hanging xlsx.js", "for (;;) {}\n", ["--ready-timeout", "2"], "no ready within 2s")
  if (haveSheetJS) await tooBigCase("E tooBig results")
  else skip("E tooBig results", `needs SheetJS at ${SHEETJS_ENTRY} (set SHEETJS_DIR)`)
  if (haveSheetJS) await maxBytesCase("F MAX_BYTES counts bytes")
  else skip("F MAX_BYTES counts bytes", `needs SheetJS at ${SHEETJS_ENTRY} (set SHEETJS_DIR)`)
  if (haveSheetJS) await timeZoneCase("G")
  else skip("G time zone", `needs SheetJS at ${SHEETJS_ENTRY} (set SHEETJS_DIR)`)
  if (haveSheetJS) await retryTimeoutCase("H timeout retried")
  else skip("H timeout retried", `needs SheetJS at ${SHEETJS_ENTRY} (set SHEETJS_DIR)`)
} finally {
  fs.rmSync(tmp, { recursive: true, force: true })
}
console.log(failed ? `${failed} check(s) failed` : "all oracle self-tests passed")
process.exit(failed ? 1 : 0)

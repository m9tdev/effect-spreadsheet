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
// Needs ~/sheetjs (SHEETJS_DIR) for A/B; Linux `ps` for the child lookup.
// Usage: node scripts/corpus/oracle-selftest.mjs

import { execFileSync, spawn } from "node:child_process"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const HERE = path.dirname(fileURLToPath(import.meta.url))
const ORACLE = path.join(HERE, "oracle.mjs")
const MANIFEST = path.join(HERE, "manifest.mjs")
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "oracle-selftest-"))
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const alive = (pid) => { try { process.kill(pid, 0); return true } catch { return false } }
const childrenOf = (pid) => {
  try { return execFileSync("ps", ["-o", "pid=", "--ppid", String(pid)], { encoding: "utf8" }).split(/\s+/).filter(Boolean).map(Number) } catch { return [] }
}

let failed = 0
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
  await killCase("A parent SIGKILL", "SIGKILL")
  await killCase("B parent SIGTERM", "SIGTERM")
  await startupCase("C throwing xlsx.js", 'throw new Error("boom from fake xlsx.js")\n', [], "boom from fake xlsx.js")
  await startupCase("D hanging xlsx.js", "for (;;) {}\n", ["--ready-timeout", "2"], "no ready within 2s")
  await tooBigCase("E tooBig results")
} finally {
  fs.rmSync(tmp, { recursive: true, force: true })
}
console.log(failed ? `${failed} check(s) failed` : "all oracle self-tests passed")
process.exit(failed ? 1 : 0)

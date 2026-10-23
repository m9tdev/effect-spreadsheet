// Tests for oracle-selftest.mjs itself. Run: node --test scripts/corpus/oracle-selftest.test.mjs
// Uses a fake SheetJS (no real one needed).
import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import test from "node:test"
import { fileURLToPath } from "node:url"

const SELFTEST = path.join(path.dirname(fileURLToPath(import.meta.url)), "oracle-selftest.mjs")

test("G2 reports the child's error line (not Node's location line) as a failed check", () => {
  const fake = fs.mkdtempSync(path.join(os.tmpdir(), "selftest-fake-sheetjs-"))
  try {
    fs.writeFileSync(path.join(fake, "xlsx.js"), 'throw new Error("boom from fake xlsx.js")\n')
    const r = spawnSync(process.execPath, [SELFTEST], { env: { ...process.env, SHEETJS_DIR: fake, ORACLE_SELFTEST_ONLY: "G" }, encoding: "utf8", timeout: 120_000 })
    assert.equal(r.status, 1, r.stdout + r.stderr) // failed checks, not a crash
    const g2 = r.stdout.split("\n").find((l) => l.includes("G2 time zone"))
    assert.ok(g2, r.stdout)
    assert.match(g2, /^FAIL .* — Error: boom from fake xlsx\.js$/)
  } finally { fs.rmSync(fake, { recursive: true, force: true }) }
})

test("a hanging G2 child is a failed check after the timeout, not a hung selftest", () => {
  const fake = fs.mkdtempSync(path.join(os.tmpdir(), "selftest-fake-sheetjs-"))
  try {
    // busy for 20 s, then it ends by itself: even if the timeout fix is reverted and this
    // test gives up, no spinning child is left behind (re-parented to init)
    fs.writeFileSync(path.join(fake, "xlsx.js"), "const t = Date.now(); while (Date.now() - t < 20_000) {}\n")
    const r = spawnSync(process.execPath, [SELFTEST], { env: { ...process.env, SHEETJS_DIR: fake, ORACLE_SELFTEST_ONLY: "G2", ORACLE_SELFTEST_G2_TIMEOUT_MS: "1000" }, encoding: "utf8", timeout: 30_000, killSignal: "SIGKILL" })
    assert.notEqual(r.signal, "SIGKILL", "the selftest itself hung")
    assert.equal(r.status, 1, r.stdout + r.stderr)
    assert.match(r.stdout, /^FAIL G2 time zone: .* — .*ETIMEDOUT/m)
  } finally { fs.rmSync(fake, { recursive: true, force: true }) }
})

test("ORACLE_SELFTEST_ONLY: a value that selects no case is an error (exit 2), not a silent pass", () => {
  for (const v of ["X", " ", ",", "G3", "a b", "A;B", "C,X"]) {
    const r = spawnSync(process.execPath, [SELFTEST], { env: { ...process.env, ORACLE_SELFTEST_ONLY: v }, encoding: "utf8", timeout: 30_000 })
    assert.equal(r.status, 2, `${JSON.stringify(v)}: ${r.stdout}${r.stderr}`)
    assert.match(r.stderr, /unknown case/, JSON.stringify(v))
    assert.doesNotMatch(r.stdout, /passed/, JSON.stringify(v))
  }
})

test("ORACLE_SELFTEST_ONLY: a partial run says so; a run where nothing ran fails", () => {
  const c = spawnSync(process.execPath, [SELFTEST], { env: { ...process.env, ORACLE_SELFTEST_ONLY: "c" }, encoding: "utf8", timeout: 60_000 })
  assert.equal(c.status, 0, c.stdout + c.stderr)
  assert.match(c.stdout, /^ok {3}C /m)
  assert.match(c.stdout, /^oracle self-tests passed \(only: C\)$/m)
  assert.doesNotMatch(c.stdout, /all oracle self-tests passed/)
  // A needs SheetJS: without it, A is only skipped, so no check runs at all
  const a = spawnSync(process.execPath, [SELFTEST], { env: { ...process.env, ORACLE_SELFTEST_ONLY: "A", SHEETJS_DIR: "/nonexistent-sheetjs" }, encoding: "utf8", timeout: 60_000 })
  assert.equal(a.status, 1, a.stdout + a.stderr)
  assert.match(a.stdout, /^no check ran \(only: A\)$/m)
})

test("ORACLE_SELFTEST_ONLY empty means the full run", () => {
  // without SheetJS the full run is quick: C and D run, the rest is skipped
  const r = spawnSync(process.execPath, [SELFTEST], { env: { ...process.env, ORACLE_SELFTEST_ONLY: "", SHEETJS_DIR: "/nonexistent-sheetjs" }, encoding: "utf8", timeout: 60_000 })
  assert.equal(r.status, 0, r.stdout + r.stderr)
  assert.match(r.stdout, /^ok {3}C /m)
  assert.match(r.stdout, /^ok {3}D /m)
  assert.match(r.stdout, /^skip H /m)
  assert.match(r.stdout, /^all oracle self-tests passed$/m)
})

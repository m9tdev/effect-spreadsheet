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

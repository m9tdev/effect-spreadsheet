// Tests for oracle.mjs's command-line behaviour. Run: node --test scripts/corpus/oracle.test.mjs
import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import test from "node:test"
import { fileURLToPath } from "node:url"

const ORACLE = path.join(path.dirname(fileURLToPath(import.meta.url)), "oracle.mjs")
const run = (extra) => {
  const cache = fs.mkdtempSync(path.join(os.tmpdir(), "oracle-test-cache-"))
  try {
    const { REQUIRE_ORACLE: _r, ...base } = process.env
    return spawnSync(process.execPath, [ORACLE, "--source", "hidden-rows", "--no-probe"], { env: { ...base, CORPUS_CACHE_DIR: cache, SHEETJS_DIR: "/nonexistent-sheetjs", ...extra }, encoding: "utf8", timeout: 30_000 })
  } finally { fs.rmSync(cache, { recursive: true, force: true }) }
}

test("missing SheetJS: exit 1 under REQUIRE_ORACLE=1, a skip (exit 0) otherwise", () => {
  const strict = run({ REQUIRE_ORACLE: "1" })
  assert.equal(strict.status, 1, strict.stdout + strict.stderr)
  assert.match(strict.stderr, /REQUIRE_ORACLE=1 but SheetJS is missing at \/nonexistent-sheetjs\/xlsx\.js/)
  for (const extra of [{}, { REQUIRE_ORACLE: "0" }]) {
    const lax = run(extra)
    assert.equal(lax.status, 0, lax.stdout + lax.stderr)
    assert.match(lax.stderr, /skipping oracle/)
  }
})

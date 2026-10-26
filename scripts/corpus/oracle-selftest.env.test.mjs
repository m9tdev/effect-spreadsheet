// Runs oracle-selftest.test.mjs with ORACLE_SELFTEST_* set in the environment, to check that
// its tests don't pass those on to the selftest. Run: node --test scripts/corpus/oracle-selftest.env.test.mjs
import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import path from "node:path"
import test from "node:test"
import { fileURLToPath } from "node:url"

const TESTS = path.join(path.dirname(fileURLToPath(import.meta.url)), "oracle-selftest.test.mjs")

test("oracle-selftest.test.mjs is not swayed by ORACLE_SELFTEST_* or REQUIRE_ORACLE in the environment", () => {
  // NODE_TEST_CONTEXT is how node:test marks its own subprocesses; drop it so the inner
  // `node --test` runs and reports on its own
  const { NODE_TEST_CONTEXT: _ctx, ...base } = process.env
  const env = { ...base, ORACLE_SELFTEST_G2_TZ: "Bogus/Zone", ORACLE_SELFTEST_ONLY: "X", ORACLE_SELFTEST_G2_TIMEOUT_MS: "1", REQUIRE_ORACLE: "1", SHEETJS_DIR: "/nonexistent-sheetjs" }
  const r = spawnSync(process.execPath, ["--test", "--test-name-pattern", "error line|partial run", TESTS], { env, encoding: "utf8", timeout: 120_000 })
  assert.match(r.stdout, /ℹ pass 2\b/, r.stdout.slice(-600)) // both selected tests ran
  assert.equal(r.status, 0, r.stdout.split("\n").filter((l) => /✖|Error/.test(l)).slice(0, 8).join("\n"))
})

// Tests for expected.mjs. Run: node --test scripts/corpus/expected.test.mjs
// Builds an oracle cache from the committed expected JSONs (no SheetJS needed).
import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import test from "node:test"
import { fileURLToPath } from "node:url"
import { CORPUS_DIR } from "./lib.mjs"

const EXPECTED = path.join(path.dirname(fileURLToPath(import.meta.url)), "expected.mjs")

function cacheFromCommitted(tamper) {
  const cache = fs.mkdtempSync(path.join(os.tmpdir(), "expected-test-cache-"))
  fs.mkdirSync(path.join(cache, "expected"))
  const root = path.join(CORPUS_DIR, "expected")
  for (const src of fs.readdirSync(root)) for (const f of fs.readdirSync(path.join(root, src))) {
    const doc = JSON.parse(fs.readFileSync(path.join(root, src, f), "utf8"))
    tamper?.(`${src}/${f.replace(/\.json$/, "")}`, doc)
    fs.writeFileSync(path.join(cache, "expected", `${doc.sha256}.json`), JSON.stringify(doc))
  }
  return cache
}

function run(cache, tmp) {
  return spawnSync(process.execPath, [EXPECTED, "--check"], { env: { ...process.env, CORPUS_CACHE_DIR: cache, TMPDIR: tmp }, encoding: "utf8" })
}

test("--check passes on a cache rebuilt from the committed files and leaves no temp dir", () => {
  const cache = cacheFromCommitted(), tmp = fs.mkdtempSync(path.join(os.tmpdir(), "expected-test-tmp-"))
  try {
    const r = run(cache, tmp)
    assert.equal(r.status, 0, r.stderr)
    assert.deepEqual(fs.readdirSync(tmp), [])
  } finally { fs.rmSync(cache, { recursive: true, force: true }); fs.rmSync(tmp, { recursive: true, force: true }) }
})

test("a failing quirk check (throws) still removes the temp dir", () => {
  const cache = cacheFromCommitted((key, doc) => { if (key === "generated/sylk_row_format_no_m.slk") doc.sheets[0].hiddenRows = [2] })
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "expected-test-tmp-"))
  try {
    const r = run(cache, tmp)
    assert.notEqual(r.status, 0)
    assert.match(r.stderr, /sylk-row-format-without-m-hidden/)
    assert.deepEqual(fs.readdirSync(tmp), [], "corpus-expected-* temp dir left behind")
  } finally { fs.rmSync(cache, { recursive: true, force: true }); fs.rmSync(tmp, { recursive: true, force: true }) }
})

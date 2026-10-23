// Tests for oracle-options.mjs. Run: node --test scripts/corpus/oracle-options.test.mjs
import assert from "node:assert/strict"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import test from "node:test"
import { SCHEMA, resultIsFresh } from "./oracle-options.mjs"

test("resultIsFresh: ok/error/tooBig results are reused; timeout, crash, old schema and junk are not", () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "oracle-fresh-"))
  try {
    const put = (name, doc) => { const p = path.join(dir, name); fs.writeFileSync(p, typeof doc === "string" ? doc : JSON.stringify(doc)); return p }
    for (const status of ["ok", "error", "tooBig"]) assert.equal(resultIsFresh(put(`${status}.json`, { schema: SCHEMA, status })), true, status)
    for (const status of ["timeout", "crash"]) assert.equal(resultIsFresh(put(`${status}.json`, { schema: SCHEMA, status })), false, status)
    assert.equal(resultIsFresh(put("old.json", { schema: SCHEMA - 1, status: "ok" })), false)
    assert.equal(resultIsFresh(put("junk.json", "{not json")), false)
    assert.equal(resultIsFresh(path.join(dir, "missing.json")), false)
  } finally { fs.rmSync(dir, { recursive: true, force: true }) }
})

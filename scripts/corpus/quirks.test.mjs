// Tests for the sheetjs-quirks.json checks. Run: node --test scripts/corpus/quirks.test.mjs
// Uses the committed fixtures and the committed expected JSONs (no SheetJS needed).
import assert from "node:assert/strict"
import fs from "node:fs"
import path from "node:path"
import test from "node:test"
import { CORPUS_DIR } from "./lib.mjs"
import { checkQuirk, csvFields, loadQuirks } from "./quirks.mjs"

const QUIRKS = loadQuirks()
const docFor = (q) => JSON.parse(fs.readFileSync(path.join(CORPUS_DIR, "expected", q.file + ".json"), "utf8"))
const fileFor = (q) => path.join(CORPUS_DIR, q.file)
const run = (q) => checkQuirk(q, docFor(q), fileFor(q))
const mutate = (q, f) => { const c = structuredClone(q); f(c); return c }

test("every quirk checks out against the oracle output and the fixture bytes", () => {
  for (const q of QUIRKS) assert.doesNotThrow(() => run(q), q.id)
})

test("every quirk has a checkable sheetjs and correct value", () => {
  for (const q of QUIRKS) {
    assert.ok(["hiddenRows", "ref", "cells", "dimension"].includes(q.field), q.id)
    assert.ok(typeof q.correctCheck === "string", `${q.id}: correctCheck`)
    if (q.field === "cells") for (const v of [q.sheetjs, q.correct]) assert.ok(Array.isArray(v) && v.every((x) => Array.isArray(x) && x.length === 3), `${q.id}: [r, c, v] triples`)
  }
})

test("a wrong sheetjs value fails, for each field type", () => {
  const byField = (f) => QUIRKS.find((q) => q.field === f)
  assert.throws(() => run(mutate(byField("hiddenRows"), (q) => { q.sheetjs = [2] })), /sheetjs: hiddenRows/)
  assert.throws(() => run(mutate(byField("cells"), (q) => { q.sheetjs[0][2] = "something else" })), /sheetjs: cells/)
  assert.throws(() => run(mutate(byField("dimension"), (q) => { q.sheetjs.ref = "A1:XFD1048576" })), /sheetjs: ref/)
  // "no cell" (v null) is checked too
  const empty = QUIRKS.find((q) => q.id === "csv-empty-quoted-field-no-cell")
  assert.throws(() => run(mutate(empty, (q) => { q.sheetjs = [[5, 0, null]] })), /sheetjs: cells/)
})

test("a wrong correct value fails, for each check method", () => {
  for (const method of new Set(QUIRKS.map((q) => q.correctCheck))) {
    const q = QUIRKS.find((x) => x.correctCheck === method)
    const wrong = mutate(q, (c) => {
      if (Array.isArray(c.correct) && Array.isArray(c.correct[0])) c.correct[0][2] = "wrong"
      else if (Array.isArray(c.correct)) c.correct = [0]
      else c.correct.declaredRows = 65536
    })
    assert.throws(() => run(wrong), /correct:/, method)
  }
})

test("csvFields: quotes, escaped quotes, explicit empty field, line endings", () => {
  const m = csvFields('a,"b ""q""",""\r\n"x\ny",\n3')
  assert.equal(m.get("0,0"), "a")
  assert.equal(m.get("0,1"), 'b "q"')
  assert.equal(m.get("0,2"), "") // "" is an explicit empty cell
  assert.equal(m.get("1,0"), "x\ny")
  assert.equal(m.has("1,1"), false) // unquoted empty is no cell
  assert.equal(m.get("2,0"), "3")
})

test("decodeFixture: BOM stripped when present, nothing lost without one", async () => {
  const { decodeFixture } = await import("./quirks.mjs")
  const be = (s) => { const b = Buffer.from(s, "utf16le"); b.swap16(); return b }
  assert.equal(decodeFixture(Buffer.concat([Buffer.from([0xfe, 0xff]), be("a,b")]), "utf-16be"), "a,b")
  assert.equal(decodeFixture(be("a,b"), "utf-16be"), "a,b") // no BOM: the first character survives
  assert.equal(decodeFixture(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from("a,b", "utf16le")]), "utf-16le"), "a,b")
  assert.equal(decodeFixture(Buffer.from("a,b", "utf16le"), "utf-16le"), "a,b")
  assert.equal(decodeFixture(Buffer.from("\ufeffa,b"), "utf-8"), "a,b")
  assert.equal(decodeFixture(Buffer.from("a,b"), "utf-8"), "a,b")
  assert.equal(decodeFixture(Buffer.from([0x80, 0x61]), "windows-1252"), "€a")
})

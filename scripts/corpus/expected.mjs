#!/usr/bin/env node
// Export committed SheetJS expectations for our own fixtures, so CI can compare
// without ~/sheetjs: test/corpus/expected/{generated,hidden-rows}/<file>.json.
//
// Reads the oracle cache (run `node scripts/corpus/oracle.mjs --source generated`
// and `--source hidden-rows` first): the parse result per sha256 plus the three
// stream.to_json probe statuses. Drops volatile fields (ms). Each file is capped at
// MAX_FILE_BYTES: over the cap, `rows` is dropped first (rows: null,
// rowsOmitted: "size cap"), then `cells` (cells: null, cellsOmitted: "size cap"),
// keeping counts, hidden rows, merges and refs. hidden-rows renamed/ copies have no file of
// their own: use hidden-rows/manifest.json `copyOf`.
//
// Usage: node scripts/corpus/expected.mjs [--check]

import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { HIDDEN_ROWS_DIR, CORPUS_DIR, EXPECTED_DIR, GENERATED_DIR, compareDirs, reportCompare, sha256 } from "./lib.mjs"

export const MAX_FILE_BYTES = 64 * 1024
const OUT_ROOT = path.join(CORPUS_DIR, "expected")
const check = process.argv.includes("--check")
const OUT = check ? fs.mkdtempSync(path.join(os.tmpdir(), "corpus-expected-")) : OUT_ROOT

function listTop(dir, skip) {
  const out = []
  const walk = (d) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name)
      const rel = path.relative(dir, p).split(path.sep).join("/")
      if (e.isDirectory()) walk(p)
      else if (!skip(rel)) out.push(rel)
    }
  }
  walk(dir)
  return out.sort()
}

const sources = [
  { name: "generated", dir: GENERATED_DIR, files: listTop(GENERATED_DIR, (r) => r === "README.json") },
  { name: "hidden-rows", dir: HIDDEN_ROWS_DIR, files: listTop(HIDDEN_ROWS_DIR, (r) => r === "manifest.json" || r.startsWith("renamed/")) },
]

// One sheet cell per line would be nicer to diff, but compact keeps the committed set small.
const stringify = (doc) => JSON.stringify(doc) + "\n"
let missing = 0, capped = 0, total = 0, count = 0
if (!check) fs.rmSync(OUT_ROOT, { recursive: true, force: true })
for (const src of sources) {
  for (const rel of src.files) {
    const sha = sha256(fs.readFileSync(path.join(src.dir, rel)))
    const cached = path.join(EXPECTED_DIR, `${sha}.json`)
    if (!fs.existsSync(cached)) { missing++; console.error(`no oracle result for ${src.name}/${rel} (run oracle.mjs --source ${src.name})`); continue }
    const doc = JSON.parse(fs.readFileSync(cached, "utf8"))
    delete doc.ms
    doc.path = rel
    doc.source = src.name
    const stream = {}
    for (const v of ["default", "cellStyles", "control"]) {
      const p = path.join(EXPECTED_DIR, `${sha}.stream.${v}.json`)
      if (fs.existsSync(p)) stream[v] = JSON.parse(fs.readFileSync(p, "utf8")).status
    }
    if (Object.keys(stream).length) doc.streamProbe = stream
    let text = stringify(doc)
    if (text.length > MAX_FILE_BYTES && doc.sheets) {
      for (const s of doc.sheets) if (s.rows) { s.rows = null; s.rowsOmitted = "size cap" }
      text = stringify(doc)
    }
    if (text.length > MAX_FILE_BYTES && doc.sheets) {
      for (const s of doc.sheets) if (s.cells) { s.cells = null; s.cellsOmitted = "size cap" }
      text = stringify(doc)
    }
    if (text.length > MAX_FILE_BYTES) throw new Error(`${src.name}/${rel}: still ${text.length} bytes after dropping rows and cells`)
    if (doc.sheets?.some((s) => s.rowsOmitted === "size cap" || s.cellsOmitted === "size cap")) capped++
    const out = path.join(OUT, src.name, rel + ".json")
    fs.mkdirSync(path.dirname(out), { recursive: true })
    fs.writeFileSync(out, text)
    total += Buffer.byteLength(text)
    count++
  }
}
console.log(`expected: ${count} files, ${(total / 1024).toFixed(1)} KiB, ${capped} capped (rows/cells dropped), ${missing} missing -> ${OUT}`)
if (check) {
  const ok = reportCompare(compareDirs(OUT, OUT_ROOT), "test/corpus/expected")
  fs.rmSync(OUT, { recursive: true, force: true })
  if (!ok || missing) process.exit(1)
} else if (missing) process.exit(1)

#!/usr/bin/env node
// Write test/corpus/hidden-rows/manifest.json: one entry per hidden-row fixture (and per
// misleading-extension copy) with the sniffed format, sha256, the rows our reader
// must return, and whether a typical SheetJS-based reader (SheetJS 0.20.3
// stream.to_json with skipHidden) hangs on it.
//
// `oldReaderHangs` was measured with SheetJS 0.20.3 (10 s kill) and is re-checked by scripts/corpus/oracle.mjs
// (probe variants "default" and "cellStyles"; see test/corpus/.cache/oracle-report.json).
//
// Usage: node scripts/corpus/hidden-rows/manifest.mjs

import fs from "node:fs"
import path from "node:path"
import { HIDDEN_ROWS_DIR, sha256, sniff } from "../lib.mjs"
import { RENAMED } from "./names.mjs"

const VALUES = [
  ["Product", "Qty", "Price"],
  ["Product A", 1, 2.5],
  ["Product B", 2, 3.75],
  ["Product C", 3, 4],
]
// Our reader never skips hidden rows: it returns all four, with hidden=true on row index 2.
const rows = (hiddenIndex) => VALUES.map((values, index) => ({ index, hidden: index === hiddenIndex, values }))

const FIXTURES = {
  "hidden-row.xls": { storage: "BIFF8 ROW record 0x0208 with fHidden 0x20 for row index 2", hidden: 2, hangs: { default: true, cellStyles: true } },
  "hidden-row.xml": { storage: '<Row ss:Index="3" ss:Hidden="1">', hidden: 2, hangs: { default: true, cellStyles: true } },
  "hidden-row.slk": { storage: "F;M0;R3 (row height 0)", hidden: 2, hangs: { default: true, cellStyles: true } },
  "hidden-row.xlsb": { storage: "BrtRowHdr flags 0x10 (fDyZero) for row index 2", hidden: 2, hangs: { default: true, cellStyles: true } },
  "hidden-row.xlsx": { storage: '<row r="3" hidden="1">', hidden: 2, hangs: { default: false, cellStyles: true }, note: "SheetJS records xlsx hidden rows in !rows only with cellStyles:true; with default options the old reader returns all 4 rows and loses the flag" },
  "plain.xlsx": { storage: "none (control)", hidden: -1, hangs: { default: false, cellStyles: false } },
}

const files = []
const entry = (rel, spec, copyOf) => {
  const buf = fs.readFileSync(path.join(HIDDEN_ROWS_DIR, rel))
  const s = sniff(buf)
  const e = { path: rel, format: s.format, size: buf.length, sha256: sha256(buf) }
  if (copyOf) {
    e.copyOf = copyOf
    e.extensionMismatch = true
    if (e.sha256 !== files.find((f) => f.path === copyOf).sha256) throw new Error(`${rel} is not a byte copy of ${copyOf}`)
  }
  e.hiddenRowStorage = spec.storage
  e.expected = { sheets: [{ name: "Sheet1", visibility: "visible", rows: rows(spec.hidden) }] }
  e.oldReaderHangs = spec.hangs
  if (spec.note) e.note = spec.note
  files.push(e)
}
for (const [rel, spec] of Object.entries(FIXTURES)) entry(rel, spec)
for (const [rel, of] of Object.entries(RENAMED)) entry(rel, FIXTURES[of], of)

const manifest = {
  version: 1,
  description: "Synthetic hidden-row fixtures (all content made up). Paths are relative to test/corpus/hidden-rows. See docs/corpus.md.",
  generator: "scripts/corpus/hidden-rows/generate.mjs (SheetJS 0.20.3 write side, script-only)",
  oldReader: "XLSX.read(buffer[, {cellStyles:true}]) then XLSX.stream.to_json(sheet, {raw:true, skipHidden:true, UTC:true, header:'A'}) (SheetJS 0.20.3)",
  hiddenRowPolicy: "Our reader keeps hidden rows and reports them with hidden=true; it never skips them.",
  files,
}
fs.writeFileSync(path.join(HIDDEN_ROWS_DIR, "manifest.json"), JSON.stringify(manifest, null, 2) + "\n")
console.log(`wrote ${files.length} entries to test/corpus/hidden-rows/manifest.json`)
for (const f of files) console.log(`  ${f.format.padEnd(5)} ${f.path.padEnd(28)} ${f.sha256}`)

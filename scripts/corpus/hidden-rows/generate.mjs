#!/usr/bin/env node
// Synthetic hidden-row fixture generator. SheetJS 0.20.3 is loaded from the local clone
// ~/sheetjs (SHEETJS_DIR), test/script-only, never from src/ or package deps.
//
// All content is made up. Document properties are set to empty values.
//
// Usage:
//   node scripts/corpus/hidden-rows/generate.mjs           # (re)write test/corpus/hidden-rows/*.{xls,xml,slk,xlsb,xlsx} and renamed/
//   node scripts/corpus/hidden-rows/generate.mjs --check   # generate to a temp dir; compare the whole directory both
//                                                 # ways with test/corpus/hidden-rows (manifest.json is hidden-rows/manifest.mjs's)
//                                                 # and the sha256s pinned in test/corpus/hidden-rows/manifest.json
//   node scripts/corpus/hidden-rows/generate.mjs --out DIR # write somewhere else

import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { createRequire } from "node:module"
import { HIDDEN_ROWS_DIR, SHEETJS_DIR, compareDirs, reportCompare, sha256 } from "../lib.mjs"
import { NAMES, RENAMED } from "./names.mjs"

const require = createRequire(import.meta.url)
const XLSX = require(path.join(SHEETJS_DIR, "xlsx.js"))

const args = process.argv.slice(2)
const check = args.includes("--check")
const outArg = args.indexOf("--out") >= 0 ? args[args.indexOf("--out") + 1] : null
const OUT = outArg ?? (check ? fs.mkdtempSync(path.join(os.tmpdir(), "hidden-rows-fixtures-")) : HIDDEN_ROWS_DIR)
fs.mkdirSync(OUT, { recursive: true })

// Row index 2 (the third row, "Product B") is the hidden row.
export const ROWS = [
  ["Product", "Qty", "Price"],
  ["Product A", 1, 2.5],
  ["Product B", 2, 3.75],
  ["Product C", 3, 4],
]
export const HIDDEN_ROW = 2

// Neutral document properties: every field SheetJS knows, empty.
const EMPTY_PROPS = { Title: "", Subject: "", Author: "", Manager: "", Company: "", Category: "", Keywords: "", Comments: "", LastAuthor: "" }

function workbook({ hidden }) {
  const ws = XLSX.utils.aoa_to_sheet(ROWS)
  if (hidden) {
    ws["!rows"] = []
    ws["!rows"][HIDDEN_ROW] = { hidden: true }
  }
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1")
  wb.Props = { ...EMPTY_PROPS }
  return wb
}

function write(name, wb, opts) {
  fs.writeFileSync(path.join(OUT, name), XLSX.write(wb, { type: "buffer", ...opts }))
}

// The SheetJS BIFF8 writer drops row properties, so the hidden flag is added by
// hand: a ROW record (0x0208) with fHidden (0x20) goes into the sheet substream,
// right after its DIMENSIONS record (0x0200).
function writeXlsWithHiddenRow(name) {
  const BOF = 0x0809, DIMENSIONS = 0x0200, ROW = 0x0208
  const cfb = XLSX.CFB.read(XLSX.write(workbook({ hidden: false }), { type: "buffer", bookType: "biff8" }), { type: "buffer" })
  const entry = XLSX.CFB.find(cfb, "Workbook")
  const records = Buffer.from(entry.content)
  let offset = 0, bofCount = 0, insertAt = -1
  // progress: offset advances by 4 + length per record
  while (offset + 4 <= records.length && insertAt < 0) {
    const type = records.readUInt16LE(offset)
    const length = records.readUInt16LE(offset + 2)
    if (type === BOF) bofCount++
    offset += 4 + length
    if (bofCount === 2 && type === DIMENSIONS) insertAt = offset
  }
  if (insertAt < 0) throw new Error("no DIMENSIONS record in the sheet")
  const row = Buffer.alloc(20)
  row.writeUInt16LE(ROW, 0) // record type
  row.writeUInt16LE(16, 2) // record length
  row.writeUInt16LE(HIDDEN_ROW, 4) // rw
  row.writeUInt16LE(0, 6) // colMic
  row.writeUInt16LE(ROWS[0].length, 8) // colMac
  row.writeUInt16LE(0xff, 10) // miyRw (height, twips)
  row.writeUInt8(0x20, 16) // fHidden
  row.writeUInt8(0x01, 17)
  row.writeUInt16LE(0x0f, 18) // ixfe
  entry.content = Buffer.concat([records.subarray(0, insertAt), row, records.subarray(insertAt)])
  entry.size = entry.content.length
  fs.writeFileSync(path.join(OUT, name), XLSX.CFB.write(cfb, { type: "buffer" }))
}

writeXlsWithHiddenRow("hidden-row.xls")
write("hidden-row.xml", workbook({ hidden: true }), { bookType: "xlml" })
write("hidden-row.slk", workbook({ hidden: true }), { bookType: "sylk" })
write("hidden-row.xlsb", workbook({ hidden: true }), { bookType: "xlsb" })
write("hidden-row.xlsx", workbook({ hidden: true }), { bookType: "xlsx" })
write("plain.xlsx", workbook({ hidden: false }), { bookType: "xlsx" })

// Misleading-extension byte copies.
for (const [rel, of] of Object.entries(RENAMED)) {
  fs.mkdirSync(path.join(OUT, path.dirname(rel)), { recursive: true })
  fs.copyFileSync(path.join(OUT, of), path.join(OUT, rel))
}
console.log(`SheetJS ${XLSX.version} (${SHEETJS_DIR}): wrote ${NAMES.length} fixtures + ${Object.keys(RENAMED).length} renamed copies to ${OUT}`)

if (check) {
  let ok = reportCompare(compareDirs(OUT, HIDDEN_ROWS_DIR, ["manifest.json"]), "test/corpus/hidden-rows")
  // The pinned sha256s in hidden-rows/manifest.json must match too (every entry, renamed copies included).
  const manifest = JSON.parse(fs.readFileSync(path.join(HIDDEN_ROWS_DIR, "manifest.json"), "utf8"))
  const listed = new Set(manifest.files.map((f) => f.path))
  for (const f of manifest.files) {
    const p = path.join(OUT, f.path)
    const got = fs.existsSync(p) ? sha256(fs.readFileSync(p)) : null
    if (got !== f.sha256) { ok = false; console.log(`PIN     ${f.path}: manifest ${f.sha256.slice(0, 12)}…, generated ${got?.slice(0, 12) ?? "missing"}`) }
  }
  for (const n of [...NAMES, ...Object.keys(RENAMED)]) if (!listed.has(n)) { ok = false; console.log(`UNLISTED ${n} (not in hidden-rows/manifest.json)`) }
  fs.rmSync(OUT, { recursive: true, force: true })
  if (!ok) process.exit(1)
  console.log("byte-identical to the committed fixtures and their manifest pins")
}

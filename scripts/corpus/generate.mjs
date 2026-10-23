#!/usr/bin/env node
// Generate small edge-case fixtures into test/corpus/generated/ (we own these; committed).
//
// Uses the SheetJS write side from ~/sheetjs (SHEETJS_DIR), test/script-only, plus
// hand-built bytes where SheetJS can't express the case (its writers drop BIFF8 row
// properties and iterate the whole declared range, so sparse/far/huge-dimension
// files are built or patched by hand). Output is deterministic: zips are rebuilt
// with fixed timestamps, document properties are fixed.
//
// Truncated / bit-flipped variants are NOT generated here: tests derive them from
// these files at test time (see docs/corpus.md).
//
// Usage: node scripts/corpus/generate.mjs [--check]
//   --check: generate into a temp dir and compare with the committed files.

import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import { createRequire } from "node:module"
import { crc32, deflateRawSync } from "node:zlib"
import { GENERATED_DIR, SHEETJS_DIR, sha256, sniff, zipEntries, zipRead } from "./lib.mjs"

const require = createRequire(import.meta.url)
const XLSX = require(path.join(SHEETJS_DIR, "xlsx.js"))

const check = process.argv.includes("--check")
const OUT = check ? fs.mkdtempSync(path.join(os.tmpdir(), "corpus-gen-")) : GENERATED_DIR
fs.rmSync(OUT, { recursive: true, force: true })
fs.mkdirSync(OUT, { recursive: true })
const written = []
const out = (name, data, desc) => {
  const p = path.join(OUT, name)
  fs.mkdirSync(path.dirname(p), { recursive: true })
  fs.writeFileSync(p, data)
  written.push({ name, size: data.length, desc })
}

// ------------------------------------------------------------------ zip writer

/** Deterministic zip: fixed DOS timestamp (1980-01-01), deflate unless tiny. */
function zip(files) {
  const locals = [], centrals = []
  let off = 0
  for (const [name, content] of files) {
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, "utf8")
    const nameBuf = Buffer.from(name, "utf8")
    const deflated = deflateRawSync(data, { level: 9 })
    const store = deflated.length >= data.length
    const body = store ? data : deflated
    const crc = crc32(data)
    const lh = Buffer.alloc(30)
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6); lh.writeUInt16LE(store ? 0 : 8, 8)
    lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0x0021, 12); lh.writeUInt32LE(crc, 14)
    lh.writeUInt32LE(body.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(nameBuf.length, 26); lh.writeUInt16LE(0, 28)
    const ch = Buffer.alloc(46)
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(store ? 0 : 8, 10)
    ch.writeUInt16LE(0, 12); ch.writeUInt16LE(0x0021, 14); ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(body.length, 20); ch.writeUInt32LE(data.length, 24)
    ch.writeUInt16LE(nameBuf.length, 28); ch.writeUInt32LE(off, 42)
    locals.push(lh, nameBuf, body)
    centrals.push(ch, nameBuf)
    off += 30 + nameBuf.length + body.length
  }
  const cd = Buffer.concat(centrals)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10)
  eocd.writeUInt32LE(cd.length, 12); eocd.writeUInt32LE(off, 16)
  return Buffer.concat([...locals, cd, eocd])
}

/** Re-zip an existing zip deterministically, optionally replacing/transforming parts. */
function rezip(buf, transform = {}) {
  const files = zipEntries(buf).filter((e) => !e.name.endsWith("/")).map((e) => {
    let data = zipRead(buf, e)
    if (transform[e.name]) data = transform[e.name](Buffer.from(data))
    return [e.name, data]
  })
  return zip(files)
}

// ------------------------------------------------------------------ SheetJS helpers

const FIXED_DATE = new Date(Date.UTC(2026, 0, 1))
const PROPS = { Title: "effect-spreadsheet fixture", Author: "effect-spreadsheet", LastAuthor: "effect-spreadsheet", Company: "", CreatedDate: FIXED_DATE, ModifiedDate: FIXED_DATE }

function book(sheets) {
  const wb = XLSX.utils.book_new()
  for (const { name, ws } of sheets) XLSX.utils.book_append_sheet(wb, ws, name)
  wb.Workbook = {}
  wb.Workbook.Sheets = sheets.map(({ name, hidden }) => ({ name, Hidden: hidden ?? 0 }))
  wb.Props = { ...PROPS }
  return wb
}
const sjs = (wb, bookType, extra = {}) => Buffer.from(XLSX.write(wb, { type: "buffer", bookType, compression: true, ...extra }))
const sjsZip = (wb, bookType, extra) => rezip(sjs(wb, bookType, extra))

// ------------------------------------------------------------------ BIFF helpers

function biffRecords(buf) {
  const recs = []
  for (let o = 0; o + 4 <= buf.length; ) {
    const id = buf.readUInt16LE(o), len = buf.readUInt16LE(o + 2)
    recs.push({ id, data: Buffer.from(buf.subarray(o + 4, o + 4 + len)) })
    o += 4 + len
  }
  return recs
}

/** Join records and fix BOUNDSHEET lbPlyPos (stream offsets of each sheet BOF). */
function biffJoin(recs) {
  const offsets = []
  let o = 0, bofs = 0
  for (const r of recs) {
    if (r.id === 0x0809 && bofs++ > 0) offsets.push(o)
    o += 4 + r.data.length
  }
  let i = 0
  for (const r of recs) if (r.id === 0x0085) r.data.writeUInt32LE(offsets[i++] ?? 0, 0)
  return Buffer.concat(recs.flatMap((r) => {
    const h = Buffer.alloc(4); h.writeUInt16LE(r.id, 0); h.writeUInt16LE(r.data.length, 2)
    return [h, r.data]
  }))
}

/** BIFF5/8 ROW record. */
function rowRecord(r, colMac, hidden) {
  const d = Buffer.alloc(16)
  d.writeUInt16LE(r, 0); d.writeUInt16LE(0, 2); d.writeUInt16LE(colMac, 4)
  d.writeUInt16LE(hidden ? 0x00ff : 0x00ff, 6)
  // flags: fDyZero (0x20) when hidden, fUnsynced (0x40), reserved byte 0x01
  d.writeUInt8(hidden ? 0x60 : 0x00, 12); d.writeUInt8(0x01, 13); d.writeUInt16LE(0x0f, 14)
  return { id: 0x0208, data: d }
}

/**
 * Patch the workbook stream of a SheetJS-written xls (BIFF5/8, in CFB).
 * plan: per sheet index { rows: [{r, colMac, hidden}], dims: {r1, c1} (DIMENSIONS last row/col, exclusive), move: [{from:[r,c], to:[r,c]}] }
 */
function patchXls(buf, plan) {
  const cfb = XLSX.CFB.read(buf, { type: "buffer" })
  const entry = XLSX.CFB.find(cfb, "Workbook") ?? XLSX.CFB.find(cfb, "Book")
  const recs = biffRecords(Buffer.from(entry.content))
  const outRecs = []
  let sheet = -1
  const CELL = new Set([0x0203, 0x00fd, 0x0204, 0x027e, 0x0201, 0x0205, 0x0006, 0x0207])
  for (const r of recs) {
    if (r.id === 0x0809) sheet++
    const p = plan[sheet - 1]
    if (p && CELL.has(r.id) && p.move) {
      for (const m of p.move) if (r.data.readUInt16LE(0) === m.from[0] && r.data.readUInt16LE(2) === m.from[1]) { r.data.writeUInt16LE(m.to[0], 0); r.data.writeUInt16LE(m.to[1], 2) }
    }
    if (p && r.id === 0x0200 && p.dims) {
      if (r.data.length >= 14) { r.data.writeUInt32LE(p.dims.r1, 4); r.data.writeUInt16LE(p.dims.c1, 10) } // BIFF8
      else { r.data.writeUInt16LE(p.dims.r1, 2); r.data.writeUInt16LE(p.dims.c1, 6) } // BIFF5
    }
    outRecs.push(r)
    if (p && r.id === 0x0200 && p.rows) for (const row of p.rows) outRecs.push(rowRecord(row.r, row.colMac, row.hidden))
  }
  entry.content = biffJoin(outRecs)
  entry.size = entry.content.length
  return Buffer.from(XLSX.CFB.write(cfb, { type: "buffer" }))
}

// ------------------------------------------------------------------ BIFF12 (xlsb) helpers

function readVar(buf, o, max) {
  let v = 0, i = 0
  for (; i < max; i++) { const b = buf[o + i]; v |= (b & 0x7f) << (7 * i); if (!(b & 0x80)) return [v >>> 0, i + 1] }
  return [v >>> 0, i]
}
function xlsbRecords(buf) {
  const recs = []
  for (let o = 0; o < buf.length; ) {
    const [type, a] = readVar(buf, o, 2)
    const [size, b] = readVar(buf, o + a, 4)
    recs.push({ type, start: o + a + b, size })
    o += a + b + size
  }
  return recs
}
/** Patch an xlsb sheet part in place: move (row,col) cells and set BrtWsDim. */
function patchXlsbSheet(buf, { move = [], dim }) {
  let curRow = -1
  for (const r of xlsbRecords(buf)) {
    if (r.type === 0x0094 && dim) { // BrtWsDim: rwFirst rwLast colFirst colLast
      buf.writeUInt32LE(dim.r0, r.start); buf.writeUInt32LE(dim.r1, r.start + 4); buf.writeUInt32LE(dim.c0, r.start + 8); buf.writeUInt32LE(dim.c1, r.start + 12)
    } else if (r.type === 0x0000) { // BrtRowHdr
      curRow = buf.readUInt32LE(r.start)
      for (const m of move) if (m.from[0] === curRow) buf.writeUInt32LE(m.to[0], r.start)
    } else if (r.type >= 1 && r.type <= 13 && r.size >= 8) { // cell records: col(4) first
      const col = buf.readUInt32LE(r.start)
      for (const m of move) if (m.from[0] === curRow && m.from[1] === col) buf.writeUInt32LE(m.to[1], r.start)
    }
  }
  return buf
}

// ------------------------------------------------------------------ hand-built xlsx

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

/**
 * sheets: [{ name, state?, dimension?, data: string (sheetData inner xml), extra?: string (after sheetData) }]
 * sst: array of raw <si> inner xml strings, numFmts: [{id, code}], xfs: [numFmtId...] (index 0 must be 0)
 */
function xlsx({ sheets, sst = null, numFmts = [], xfs = [0], date1904 = false }) {
  const files = []
  const ov = sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")
  files.push(["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${ov}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sst ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : ""}</Types>`])
  files.push(["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`])
  const sh = sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ""} r:id="rId${i + 1}"/>`).join("")
  files.push(["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook ${NS}>${date1904 ? '<workbookPr date1904="1"/>' : "<workbookPr/>"}<sheets>${sh}</sheets></workbook>`])
  const n = sheets.length
  const rels = sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")
    + `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
    + (sst ? `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` : "")
  files.push(["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`])
  const nf = numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((f) => `<numFmt numFmtId="${f.id}" formatCode="${esc(f.code)}"/>`).join("")}</numFmts>` : ""
  const xf = xfs.map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"${id ? ' applyNumberFormat="1"' : ""}/>`).join("")
  files.push(["xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${nf}<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xf}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`])
  if (sst) files.push(["xl/sharedStrings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.length}" uniqueCount="${sst.length}">${sst.map((s) => `<si>${s}</si>`).join("")}</sst>`])
  sheets.forEach((s, i) => files.push([`xl/worksheets/sheet${i + 1}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ${NS}>${s.dimension ? `<dimension ref="${s.dimension}"/>` : ""}<sheetData>${s.data}</sheetData>${s.extra ?? ""}</worksheet>`]))
  return zip(files)
}

// ------------------------------------------------------------------ hand-built SpreadsheetML 2003

function xlml(worksheets) {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n${worksheets.join("\n")}\n</Workbook>\n`
}

// ================================================================== fixtures

// ---- 1. hidden rows: first, middle, last, consecutive, all (6 rows x 2 cols)
const HR_ROWS = Array.from({ length: 6 }, (_, i) => [`r${i + 1}`, i + 1])
const HR_CASES = { first: [0], middle: [2], last: [5], consecutive: [1, 2, 3], all: [0, 1, 2, 3, 4, 5] }
const hrSheet = (hidden) => {
  const ws = XLSX.utils.aoa_to_sheet(HR_ROWS)
  ws["!rows"] = []
  for (const r of hidden) ws["!rows"][r] = { hidden: true }
  return ws
}
const hrBook = () => book(Object.entries(HR_CASES).map(([name, hidden]) => ({ name, ws: hrSheet(hidden) })))
const hrDesc = "hidden rows; sheets first/middle/last/consecutive/all, 6 rows x 2 cols (A=r1..r6, B=1..6)"
out("hidden_rows.xlsx", sjsZip(hrBook(), "xlsx"), hrDesc)
out("hidden_rows.xlsb", sjsZip(hrBook(), "xlsb"), hrDesc)
out("hidden_rows.xml", Buffer.from(XLSX.write(hrBook(), { type: "string", bookType: "xlml" }), "utf8"), hrDesc + " (SpreadsheetML 2003)")
{
  // SheetJS's BIFF writers drop row properties: insert ROW records by hand.
  const plan = Object.values(HR_CASES).map((hidden) => ({ rows: HR_ROWS.map((_, r) => ({ r, colMac: 2, hidden: hidden.includes(r) })) }))
  const plain = book(Object.keys(HR_CASES).map((name) => ({ name, ws: XLSX.utils.aoa_to_sheet(HR_ROWS) })))
  out("hidden_rows.xls", patchXls(sjs(plain, "biff8"), plan), hrDesc + " (BIFF8, ROW records hand-inserted)")
  out("hidden_rows.biff5.xls", patchXls(sjs(plain, "biff5"), plan), hrDesc + " (BIFF5, ROW records hand-inserted)")
}
for (const [name, hidden] of Object.entries(HR_CASES)) {
  out(`hidden_rows_${name}.slk`, Buffer.from(XLSX.write(book([{ name: "Sheet1", ws: hrSheet(hidden) }]), { type: "string", bookType: "sylk" }), "latin1"), `SYLK hidden rows (${name}): F;M0;R<n>`)
}
{
  // ODS: SheetJS doesn't write row visibility; add table:visibility="collapse" by hand.
  const ods = sjs(hrBook(), "ods")
  const patched = rezip(ods, {
    "content.xml": (b) => {
      const cases = Object.values(HR_CASES)
      let t = -1, row = 0
      const s = b.toString("utf8").replace(/<table:table [^>]*>|<table:table-row\b/g, (m) => {
        if (m.startsWith("<table:table ")) { t++; row = 0; return m }
        const r = row++
        return cases[t]?.includes(r) ? '<table:table-row table:visibility="collapse"' : m
      })
      return Buffer.from(s, "utf8")
    },
  })
  out("hidden_rows.ods", patched, hrDesc + ' (ODS, table:visibility="collapse" hand-added)')
}

// ---- 2. sparse / far cells: A1 and XFD1048576 (format max)
{
  out("sparse_far.xlsx", xlsx({ sheets: [{ name: "Sheet1", dimension: "A1:XFD1048576", data: '<row r="1"><c r="A1" t="inlineStr"><is><t>first</t></is></c></row><row r="1048576"><c r="XFD1048576"><v>42</v></c></row>' }] }), "sparse/far: A1='first', XFD1048576=42")
  const xb = sjs(book([{ name: "Sheet1", ws: XLSX.utils.aoa_to_sheet([["first", null], [null, 42]]) }]), "xlsb")
  out("sparse_far.xlsb", rezip(xb, { "xl/worksheets/sheet1.bin": (b) => patchXlsbSheet(b, { move: [{ from: [1, 1], to: [1048575, 16383] }], dim: { r0: 0, r1: 1048575, c0: 0, c1: 16383 } }) }), "sparse/far: A1='first', XFD1048576=42 (BIFF12 patched)")
  const xl = sjs(book([{ name: "Sheet1", ws: XLSX.utils.aoa_to_sheet([["first", null], [null, 42]]) }]), "biff8")
  out("sparse_far.xls", patchXls(xl, [{ move: [{ from: [1, 1], to: [65535, 255] }], dims: { r1: 65536, c1: 256 } }]), "sparse/far: A1='first', IV65536=42 (BIFF8 max)")
  out("sparse_far.xml", Buffer.from(xlml([`<Worksheet ss:Name="Sheet1"><Table><Row ss:Index="1"><Cell ss:Index="1"><Data ss:Type="String">first</Data></Cell></Row><Row ss:Index="1048576"><Cell ss:Index="16384"><Data ss:Type="Number">42</Data></Cell></Row></Table></Worksheet>`])), "sparse/far: A1='first', XFD1048576=42")
  out("sparse_far.slk", Buffer.from('ID;PWXL;N;E\r\nC;Y1;X1;K"first"\r\nC;Y1048576;X16384;K42\r\nE\r\n', "latin1"), "sparse/far: A1='first', XFD1048576=42")
}

// ---- 3. huge declared dimension with only 2 cells (A1, B2)
{
  const two = '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row>'
  out("huge_dimension.xlsx", xlsx({ sheets: [{ name: "Sheet1", dimension: "A1:XFD1048576", data: two }] }), 'huge <dimension ref="A1:XFD1048576"/>, cells A1=1, B2=2 only')
  const ws = XLSX.utils.aoa_to_sheet([[1, null], [null, 2]])
  out("huge_dimension.xlsb", rezip(sjs(book([{ name: "Sheet1", ws }]), "xlsb"), { "xl/worksheets/sheet1.bin": (b) => patchXlsbSheet(b, { dim: { r0: 0, r1: 1048575, c0: 0, c1: 16383 } }) }), "huge BrtWsDim A1:XFD1048576, cells A1=1, B2=2 only")
  out("huge_dimension.xls", patchXls(sjs(book([{ name: "Sheet1", ws: XLSX.utils.aoa_to_sheet([[1, null], [null, 2]]) }]), "biff8"), [{ dims: { r1: 65536, c1: 256 } }]), "huge DIMENSIONS A1:IV65536, cells A1=1, B2=2 only")
  out("huge_dimension.xml", Buffer.from(xlml([`<Worksheet ss:Name="Sheet1"><Table ss:ExpandedColumnCount="16384" ss:ExpandedRowCount="1048576"><Row><Cell><Data ss:Type="Number">1</Data></Cell></Row><Row><Cell ss:Index="2"><Data ss:Type="Number">2</Data></Cell></Row></Table></Worksheet>`])), "huge ExpandedRowCount/ColumnCount, cells A1=1, B2=2 only")
  out("huge_dimension.slk", Buffer.from("ID;PWXL;N;E\r\nB;Y1048576;X16384;D0 0 1048575 16383\r\nC;Y1;X1;K1\r\nC;Y2;X2;K2\r\nE\r\n", "latin1"), "huge B;Y1048576;X16384, cells A1=1, B2=2 only")
}

// ---- 4. many empty rows: 100k empty <row> elements between two cells (some hidden, some with height)
{
  const parts = ['<row r="1"><c r="A1" t="inlineStr"><is><t>top</t></is></c></row>']
  for (let r = 2; r <= 5001; r++) parts.push(r % 1000 === 0 ? `<row r="${r}" hidden="1"/>` : r % 777 === 0 ? `<row r="${r}" ht="30" customHeight="1"/>` : `<row r="${r}"/>`)
  parts.push('<row r="5002"><c r="A5002" t="inlineStr"><is><t>bottom</t></is></c></row>')
  out("many_empty_rows.xlsx", xlsx({ sheets: [{ name: "Sheet1", dimension: "A1:A5002", data: parts.join("") }] }), "5000 empty <row r=.../> elements between A1='top' and A5002='bottom'; every 1000th empty row hidden, every 777th has a custom height")
  // 1048574 empty rows without r attributes (implicit numbering) then a cell on the last row: ~6 MB of XML, a few KB zipped
  const implicit = '<row><c t="inlineStr"><is><t>top</t></is></c></row>' + "<row/>".repeat(1048574) + '<row><c t="inlineStr"><is><t>bottom</t></is></c></row>'
  out("many_empty_rows_implicit.xlsx", xlsx({ sheets: [{ name: "Sheet1", data: implicit }] }), "1048574 empty <row/> without r (implicit numbering) between A1='top' and A1048576='bottom'; ~6 MB XML, high zip ratio")
  out("many_empty_rows.csv", Buffer.from("top\n" + "\n".repeat(5000) + "bottom\n"), "5000 blank lines between 'top' and 'bottom'")
}

// ---- 5. strings: shared (plain, rich, entities, _xHHHH_, xml:space), inline, str formula, implicit refs
{
  const sst = [
    "<t>plain shared</t>",
    '<r><t>rich </t></r><r><rPr><b/></rPr><t>bold</t></r><r><t xml:space="preserve"> tail </t></r>',
    "<t>&lt;tag&gt; &amp; &quot;quotes&quot; &#x1F363; &#233;</t>",
    "<t>line1_x000D__x000A_line2 _x0041_</t>",
    '<t xml:space="preserve">  leading and trailing  </t>',
    "<t>日本語 العربية עברית Ελληνικά</t>",
    "<t/>",
    '<t>phonetic</t><rPh sb="0" eb="1"><t>ignored</t></rPh>',
  ]
  const data = [
    '<row r="1">' + sst.map((_, i) => `<c r="${XLSX.utils.encode_col(i)}1" t="s"><v>${i}</v></c>`).join("") + "</row>",
    '<row r="2"><c r="A2" t="inlineStr"><is><t>inline</t></is></c><c r="B2" t="inlineStr"><is><r><t>inline </t></r><r><rPr><i/></rPr><t>rich</t></r></is></c><c r="C2" t="str"><f>"str"&amp;"ing"</f><v>string</v></c><c r="D2" t="b"><v>1</v></c><c r="E2" t="e"><v>#N/A</v></c><c r="F2" t="d"><v>2026-01-02T03:04:05Z</v></c><c r="G2" t="n"><v>-1.5E-10</v></c></row>',
    // implicit references: cells and rows without r attributes
    "<row><c><v>10</v></c><c><v>11</v></c><c r=\"E3\"><v>14</v></c><c><v>15</v></c></row>",
    '<row><c t="inlineStr"><is><t>implicit row 4</t></is></c></row>',
  ].join("")
  out("strings.xlsx", xlsx({ sheets: [{ name: "Strings", dimension: "A1:H4", data }], sst }), "shared strings (rich runs, entities, _xHHHH_ escapes, xml:space, unicode, empty, phonetic), inline strings, t=str/b/e/d cells, rows/cells without r attributes")
}

// ---- 6. formulas with cached values (plain, shared, array), xlsx hand-built
{
  const data = [
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c><c r="C1"><f t="shared" ref="C1:C3" si="0">A1+10</f><v>11</v></c><c r="D1" t="str"><f>IF(A1&gt;0,"pos","neg")</f><v>pos</v></c></row>',
    '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f>A2*2</f><v>4</v></c><c r="C2"><f t="shared" si="0"/><v>12</v></c><c r="D2" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>',
    '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="array" ref="B3">SUM(A1:A3*1)</f><v>6</v></c><c r="C3"><f t="shared" si="0"/><v>13</v></c><c r="D3" t="b"><f>A3=3</f><v>1</v></c><c r="E3"><f>NOW()</f></c></row>',
  ].join("")
  out("formulas.xlsx", xlsx({ sheets: [{ name: "Formulas", dimension: "A1:E3", data }] }), "formulas with cached values: plain, shared (si), array, string/error/bool results, one formula without cached value")
}

// ---- 7. dates with number formats (xlsx hand-built + 1904 variant)
{
  const fmts = [
    { id: 14 }, { id: 15 }, { id: 17 }, { id: 20 }, { id: 22 }, { id: 45 }, { id: 46 }, { id: 47 },
    { id: 164, code: "yyyy-mm-dd" }, { id: 165, code: "dd/mm/yyyy hh:mm:ss.000" }, { id: 166, code: "[h]:mm:ss" },
    { id: 167, code: '[$-409]mmmm d, yyyy;@' }, { id: 168, code: '"Date:" yyyy' }, { id: 169, code: "0.00" }, { id: 170, code: '[Red][<0]"neg";0' }, { id: 171, code: 'h "h" mm "m"' }, { id: 172, code: "yyyy\\-mm" },
  ]
  const numFmts = fmts.filter((f) => f.code)
  const xfs = [0, ...fmts.map((f) => f.id)]
  const serial = 46024.5208333333 // 2026-01-02 12:30 (1900 system)
  const row = (r, v) => `<row r="${r}">` + fmts.map((f, i) => `<c r="${XLSX.utils.encode_col(i)}${r}" s="${i + 1}"><v>${v}</v></c>`).join("") + "</row>"
  const data = row(1, serial) + row(2, 0) + row(3, 60) + row(4, 61) + row(5, 1.75) + row(6, -1)
  const sheets = [{ name: "Dates", dimension: `A1:${XLSX.utils.encode_col(fmts.length - 1)}6`, data }]
  out("dates.xlsx", xlsx({ sheets, numFmts, xfs }), "date/time number formats: builtins 14,15,17,20,22,45-47 and custom (yyyy-mm-dd, [h]:mm:ss, locale prefix, quoted text, non-date 0.00, ...); values 46024.52, 0, 60 (1900-02-29 bug), 61, 1.75, -1")
  out("dates_1904.xlsx", xlsx({ sheets, numFmts, xfs, date1904: true }), "same as dates.xlsx with workbookPr date1904=1")
}

// ---- 8. features workbook via SheetJS, in every writable format
{
  const data = XLSX.utils.aoa_to_sheet([
    ["text", "number", "bool", "unicode", "formula"],
    ["héllo", 1, true, "日本語", null],
    ["ünïcödé 🍣", -2.5, false, "עברית", null],
    ["", 1e308, null, "Ελληνικά", null],
  ])
  data["E2"] = { t: "n", v: 2, f: "B2*2" }
  data["E3"] = { t: "n", v: -5, f: "B3*2" }
  data["E4"] = { t: "s", v: "x", f: "\"x\"" }
  data["F2"] = { t: "e", v: 0x2a } // #N/A
  data["F3"] = { t: "e", v: 0x07 } // #DIV/0!
  data["!ref"] = "A1:F4"
  const dates = XLSX.utils.aoa_to_sheet([["fmt", "value"]])
  const dfmts = ["yyyy-mm-dd", "m/d/yy", "d-mmm-yy", "h:mm:ss", "[h]:mm", "yyyy-mm-dd hh:mm:ss", "dd/mm/yyyy", "0.00"]
  dfmts.forEach((z, i) => { dates[`A${i + 2}`] = { t: "s", v: z }; dates[`B${i + 2}`] = { t: "n", v: 46024.5208333333, z } })
  dates["!ref"] = `A1:B${dfmts.length + 1}`
  const merges = XLSX.utils.aoa_to_sheet([["h-merge", null, null, "x"], ["v-merge", "rect", null, null], [null, null, null, null], [null, null, null, "end"]])
  merges["!merges"] = [XLSX.utils.decode_range("A1:C1"), XLSX.utils.decode_range("A2:A3"), XLSX.utils.decode_range("B2:C3")]
  const fb = () => book([
    { name: "Data", ws: data },
    { name: "Dates", ws: dates },
    { name: "Merges", ws: merges },
    { name: "Hidden", ws: XLSX.utils.aoa_to_sheet([["hidden sheet"]]), hidden: 1 },
    { name: "VeryHidden", ws: XLSX.utils.aoa_to_sheet([["very hidden sheet"]]), hidden: 2 },
    { name: "Ünïcödé 名前", ws: XLSX.utils.aoa_to_sheet([["sheet name with unicode"]]) },
  ])
  const fdesc = "sheets Data (text/number/bool/unicode/formulas with cached values/errors), Dates (8 formats), Merges (A1:C1, A2:A3, B2:C3), Hidden, VeryHidden, unicode sheet name"
  out("features.xlsx", sjsZip(fb(), "xlsx", { bookSST: true }), fdesc + " (shared strings)")
  out("features.xlsb", sjsZip(fb(), "xlsb"), fdesc)
  out("features.xls", sjs(fb(), "biff8"), fdesc)
  out("features.biff5.xls", sjs(fb(), "biff5"), fdesc + " (BIFF5)")
  out("features.xml", Buffer.from(XLSX.write(fb(), { type: "string", bookType: "xlml" }), "utf8"), fdesc + " (SpreadsheetML 2003)")
  const quiet = (f) => { const e = console.error; console.error = () => {}; try { return f() } finally { console.error = e } } // SheetJS warns about ODF number formats
  out("features.ods", quiet(() => sjsZip(fb(), "ods")), fdesc + " (ODS)")
  out("features.fods", quiet(() => Buffer.from(XLSX.write(fb(), { type: "string", bookType: "fods" }), "utf8")), fdesc + " (flat ODS)")
  out("features.slk", Buffer.from(XLSX.write(fb(), { type: "string", bookType: "sylk" }), "latin1"), "first sheet (Data) of the features workbook as SYLK")
  const small = () => book([{ name: "Sheet1", ws: XLSX.utils.aoa_to_sheet([["a", 1], ["b", 2.5], [true, "c"]]) }])
  out("basic.biff2.xls", sjs(small(), "biff2"), "BIFF2 raw stream: a,1 / b,2.5 / TRUE,c")
  out("basic.biff3.xls", sjs(small(), "biff3"), "BIFF3 raw stream: a,1 / b,2.5 / TRUE,c")
  out("basic.biff4.xls", sjs(small(), "biff4"), "BIFF4 raw stream: a,1 / b,2.5 / TRUE,c")
}

// ---- 9. CSV variants
{
  const utf16le = (s) => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(s, "utf16le")])
  const utf16be = (s) => { const b = Buffer.from(s, "utf16le"); b.swap16(); return Buffer.concat([Buffer.from([0xfe, 0xff]), b]) }
  out("csv_bom_utf8.csv", Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("name,city\r\nJosé,Zürich\r\n名前,東京\r\n")]), "UTF-8 with BOM, CRLF")
  out("csv_semicolon.csv", Buffer.from('name;amount;date\r\n"Smith; J.";1.234,56;31.12.2025\r\nDoe;-0,5;01.01.2026\r\n'), "; delimiter, decimal comma (European Excel export)")
  out("csv_utf16le_tab.txt", utf16le("name\tvalue\r\nalpha\t1\r\nβeta\t2\r\n🍣\t3\r\n"), "UTF-16LE with BOM, tab-delimited (Excel 'Unicode Text')")
  out("csv_utf16be.csv", utf16be("a,b\r\n1,é\r\n"), "UTF-16BE with BOM")
  out("csv_quoted_newlines.csv", Buffer.from('id,text\n1,"line one\nline two"\n2,"crlf\r\ninside"\n3,"say ""hi"""\n4,"comma, inside"\n5,""\n'), 'quoted LF and CRLF inside fields, escaped "" quotes, empty quoted field')
  out("csv_cr_only.csv", Buffer.from("a,b\r1,2\r3,4\r"), "old Mac CR-only line endings")
  out("csv_windows1252.csv", Buffer.from([...Buffer.from("name,price\r\nCaf"), 0xe9, ...Buffer.from(",5\r\n"), 0x80, ...Buffer.from("uro,"), 0x93, ...Buffer.from("q"), 0x94, ...Buffer.from("\r\n")]), "Windows-1252 bytes (é, €, curly quotes), no BOM")
  out("csv_ragged.csv", Buffer.from("a,b,c\n1\n1,2,3,4,5\n,,\n"), "ragged rows (1..5 fields), trailing empty fields")
  out("csv_no_trailing_newline.csv", Buffer.from("a,b\n1,2"), "no newline at end of file")
  out("csv_empty.csv", Buffer.alloc(0), "empty file (0 bytes)")
  out("csv_tab.tsv", Buffer.from("a\tb\tc\n1\t\t3\n"), "tab-delimited, UTF-8, empty middle field")
  out("csv_bom_only.csv", Buffer.from([0xef, 0xbb, 0xbf]), "UTF-8 BOM and nothing else")
}

// ---- index
const index = { generator: "scripts/corpus/generate.mjs", sheetjs: XLSX.version, files: written.map((w) => {
  const buf = fs.readFileSync(path.join(OUT, w.name))
  return { name: w.name, format: sniff(buf).format, size: w.size, sha256: sha256(buf), desc: w.desc }
}) }
fs.writeFileSync(path.join(OUT, "README.json"), JSON.stringify(index, null, 2) + "\n")
const total = written.reduce((a, w) => a + w.size, 0)
console.log(`generated ${written.length} files, ${(total / 1024).toFixed(1)} KiB total -> ${OUT}`)

if (check) {
  let bad = 0
  for (const w of index.files) {
    const p = path.join(GENERATED_DIR, w.name)
    const same = fs.existsSync(p) && sha256(fs.readFileSync(p)) === w.sha256
    if (!same) { bad++; console.log(`DIFF ${w.name}`) }
  }
  fs.rmSync(OUT, { recursive: true, force: true })
  if (bad) process.exit(1)
  console.log("all generated files byte-identical to the committed ones")
}

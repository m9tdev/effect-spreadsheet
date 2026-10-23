// Checks for test/corpus/sheetjs-quirks.json (value formats are documented in that file):
// each quirk's `sheetjs` value against the oracle result, and its `correct` value
// against the fixture bytes (method `correctCheck`). Used by expected.mjs; tested by
// quirks.test.mjs.

import fs from "node:fs"
import path from "node:path"
import { CORPUS_DIR, sniff, zipEntries, zipRead } from "./lib.mjs"

export const QUIRKS_PATH = path.join(CORPUS_DIR, "sheetjs-quirks.json")
export const loadQuirks = () => JSON.parse(fs.readFileSync(QUIRKS_PATH, "utf8")).quirks

const colName = (c) => { let s = ""; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s }
const a1 = (r, c) => `${colName(c)}${r + 1}`
const sameCells = (got, want) => want.every(([r, c, v]) => (v === null ? !got.has(`${r},${c}`) : got.get(`${r},${c}`) === v))

// Minimal RFC 4180 reader for the quirk checks (quotes, "" escapes, CRLF/LF/CR rows).
export function csvFields(text) {
  const rows = [[]]
  let f = "", q = false, any = false
  // progress: i increases every step
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++ } else q = false } else f += ch; continue }
    if (ch === '"') { q = true; any = true }
    else if (ch === ",") { rows.at(-1).push({ f, any }); f = ""; any = false }
    else if (ch === "\r" || ch === "\n") { if (ch === "\r" && text[i + 1] === "\n") i++; rows.at(-1).push({ f, any }); rows.push([]); f = ""; any = false }
    else f += ch
  }
  if (f || any || rows.at(-1).length) rows.at(-1).push({ f, any })
  if (rows.at(-1).length === 0) rows.pop()
  const m = new Map()
  // a field is a cell when it has text or was quoted ("" is an explicit empty string)
  rows.forEach((row, r) => row.forEach((x, c) => { if (x.f !== "" || x.any) m.set(`${r},${c}`, x.f) }))
  return m
}
// TextDecoder strips a leading BOM itself for utf-8 and utf-16 (ignoreBOM is false by
// default), so nothing is skipped here: a BOM-less file keeps its first character.
export function decodeFixture(buf, enc) {
  return new TextDecoder(enc).decode(buf)
}
export const xmlText = (s) => s
  .replace(/<rPh\b[\s\S]*?<\/rPh>/g, "")
  .replace(/<t(?:\s[^>]*)?\/>/g, "")
  .replace(/[\s\S]*?<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, "$1\u0000").split("\u0000").slice(0, -1).join("")
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d))
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&")
  .replace(/_x([0-9A-F]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
function xlsxSharedStringCells(buf) {
  const part = (n) => { const e = zipEntries(buf).find((x) => x.name === n); return e ? zipRead(buf, e).toString("utf8") : "" }
  const sst = [...part("xl/sharedStrings.xml").matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => xmlText(m[1]))
  const m = new Map()
  for (const c of part("xl/worksheets/sheet1.xml").matchAll(/<c r="([A-Z]+)(\d+)" t="s"><v>(\d+)<\/v><\/c>/g)) {
    const col = [...c[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    m.set(`${+c[2] - 1},${col}`, sst[+c[3]])
  }
  return m
}
export function checkQuirk(q, doc, file) {
  const bad = (why) => { throw new Error(`sheetjs-quirks.json ${q.id} (${q.file}): ${why}`) }
  const sh = doc.sheets?.find((x) => x.name === q.sheet)
  if (!sh) bad(`no sheet ${q.sheet}`)
  const oracleCells = new Map((sh.cells ?? []).map((x) => [`${x.r},${x.c}`, x.v]))
  // sheetjs value vs the oracle
  if (q.field === "hiddenRows") { if (JSON.stringify(sh.hiddenRows) !== JSON.stringify(q.sheetjs)) bad(`sheetjs: hiddenRows ${JSON.stringify(sh.hiddenRows)}`) }
  else if (q.field === "ref") { if (sh.ref !== q.sheetjs) bad(`sheetjs: ref ${sh.ref}`) }
  else if (q.field === "dimension") { if (sh.ref !== q.sheetjs.ref) bad(`sheetjs: ref ${sh.ref}`) }
  else if (q.field === "rows") {
    // sheet_to_json rows (header:1, from the !ref's first row/col); Dates appear as ISO strings
    const rowCells = new Map()
    ;(sh.rows ?? []).forEach((row, r) => (row ?? []).forEach((v, c) => { if (v !== null) rowCells.set(`${r},${c}`, v) }))
    if (!Array.isArray(q.sheetjs) || !sameCells(rowCells, q.sheetjs)) bad(`sheetjs: rows ${JSON.stringify(q.sheetjs?.map?.(([r, c]) => [r, c, rowCells.get(`${r},${c}`) ?? null]))}`)
  }
  else if (q.field === "cells") { if (!Array.isArray(q.sheetjs) || !sameCells(oracleCells, q.sheetjs)) bad(`sheetjs: cells ${JSON.stringify(q.sheetjs?.map?.(([r, c]) => [r, c, oracleCells.get(`${r},${c}`) ?? null]))}`) }
  else bad(`unknown field ${q.field}`)
  // correct value vs the fixture bytes
  const buf = fs.readFileSync(file)
  const m = q.correctCheck
  if (m === "sylk-hidden-rows") {
    const rows = [...buf.toString("latin1").matchAll(/^F;([^\r\n]*)$/gm)].map((x) => x[1].split(";")).filter((f) => f.includes("M0")).flatMap((f) => f.filter((x) => /^R\d+$/.test(x)).map((x) => +x.slice(1) - 1))
    if (JSON.stringify(rows.sort((a, b) => a - b)) !== JSON.stringify(q.correct)) bad(`correct: M0 rows in the file are ${JSON.stringify(rows)}`)
  } else if (m?.startsWith("csv:")) {
    const got = csvFields(decodeFixture(buf, m.slice(4)))
    if (!sameCells(got, q.correct)) bad(`correct: decoded fields ${JSON.stringify(q.correct.map(([r, c]) => [r, c, got.get(`${r},${c}`) ?? null]))}`)
  } else if (m === "xlsx-shared-strings") {
    const got = xlsxSharedStringCells(buf)
    if (!sameCells(got, q.correct)) bad(`correct: shared strings ${JSON.stringify(q.correct.map(([r, c]) => [r, c, got.get(`${r},${c}`) ?? null]))}`)
  } else if (m === "xlsx-date1904-serial") {
    // the serials in the sheet XML read in the workbook's own date system (date1904 => 1904-01-01 epoch)
    const part = (n) => { const e = zipEntries(buf).find((x) => x.name === n); return e ? zipRead(buf, e).toString("utf8") : "" }
    const d1904 = /<workbookPr\b[^>]*\bdate1904="(1|true)"/.test(part("xl/workbook.xml"))
    if (!d1904) bad("correct: the workbook doesn't declare date1904")
    const epoch = Date.UTC(1904, 0, 1)
    const got = new Map()
    for (const c of part("xl/worksheets/sheet1.xml").matchAll(/<c r="([A-Z]+)(\d+)"([^>]*)><v>([^<]+)<\/v><\/c>/g)) {
      // only plain numeric cells are serials: skip typed cells (t="s", "str", "b", ...) and
      // values that aren't finite numbers, which would make new Date() throw a RangeError
      if (/\bt="/.test(c[3])) continue
      const serial = Number(c[4])
      if (!Number.isFinite(serial)) continue
      const ms = epoch + Math.round(serial * 86400) * 1000
      if (!Number.isFinite(new Date(ms).getTime())) continue
      const col = [...c[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
      got.set(`${+c[2] - 1},${col}`, new Date(ms).toISOString())
    }
    if (!sameCells(got, q.correct)) bad(`correct: 1904-system dates ${JSON.stringify(q.correct.map(([r, c]) => [r, c, got.get(`${r},${c}`) ?? null]))}`)
  } else if (m === "declared-dimension") {
    const rs = (sh.cells ?? []).map((x) => x.r), cs = (sh.cells ?? []).map((x) => x.c)
    const range = `${a1(Math.min(...rs), Math.min(...cs))}:${a1(Math.max(...rs), Math.max(...cs))}`
    const hd = /rows (\d+), cols (\d+)/.exec(sniff(buf).features.hugeDimension ?? "")
    if (range !== q.correct.cells) bad(`correct: cells span ${range}`)
    if (!hd || +hd[1] !== q.correct.declaredRows || +hd[2] !== q.correct.declaredCols) bad(`correct: declared ${hd?.[0] ?? "none"}`)
  } else bad(`unknown correctCheck ${m}`)
}


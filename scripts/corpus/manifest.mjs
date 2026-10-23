#!/usr/bin/env node
// Build the full manifest test/corpus/.cache/manifest.json (gitignored: it lists
// third-party SheetJS corpus files, which we don't publish) from the fetched
// SheetJS corpus, our generated fixtures and the hidden-row fixtures. Formats are sniffed from bytes (scripts/corpus/lib.mjs),
// never from the file extension.
//
// Usage: node scripts/corpus/manifest.mjs [--all-formats]
//   default: lists files in the focus formats only (FOCUS below);
//   --all-formats: lists every data file (incl. lotus, dif, html, json, ...).

import fs from "node:fs"
import path from "node:path"
import { HIDDEN_ROWS_DIR, CACHE_DIR, GENERATED_DIR, MANIFEST_PATH, SHEETJS_CACHE, SHEETJS_DIR, parseFlags, sha256, sniff } from "./lib.mjs"

export const FOCUS = new Set(["csv", "xlsx", "xlsm", "xltx", "xls", "xlsb", "xlml", "ods", "fods", "sylk", "ooxml-encrypted"])

const allFormats = !!parseFlags(process.argv.slice(2), ["--all-formats"])["--all-formats"]
const t0 = Date.now()
const SHEETJS_ROOT = path.join(SHEETJS_CACHE, "test_files")

// Tooling/docs in test_files that are not corpus data (excluded by name, not sniffed).
const NON_DATA = /(^|\/)(README\.md|LICENSE|Makefile|package\.json|requirements\.txt|\.rooster\.yaml|\.travis\.yml|\.gitignore|tests\.skip|\.DS_Store|\.fetched\.json)$|\.(sheetnames|sh|scpt|rb|py|md)$|^(tests|misc)\//

// Provenance by prefix, per test_files/README.md: a derivative of an external
// file carries the source prefix (`<prefix>_`) in its name; external clones live
// in `<owner>_<repo>/` directories. SheetJS-original (Apache-2.0) is only claimed for
// root files and the Excel re-saves in 201*/ and biff*/ whose name has no prefix or a
// stem documented in the test_files README (SHEETJS_STEMS). Anything else is "unclear":
// other subdirectories, and names with an unknown `<prefix>_`.
// `committable` is our conclusion (docs/corpus.md); we still commit none of them for now.
const ORIGINS = [
  { re: /^apachepoi_|^apache_spreadsheet\/|^apache_ExcelXML\//, origin: "apache-poi", license: "Apache-2.0", committable: true },
  { re: /^spout-(xlsx|ods)_/, origin: "spout", license: "Apache-2.0", committable: true },
  { re: /^oo34xml_/, origin: "openoffice", license: "Apache-2.0", committable: true },
  { re: /^openpyxl_/, origin: "openpyxl", license: "MIT", committable: true },
  { re: /^roo_|^Empact_roo\/|^roo-rb_roo-xls\//, origin: "roo", license: "MIT", committable: true },
  { re: /^spreadsheet-parsexlsx_|^doy_spreadsheet-parsexlsx\//, origin: "spreadsheet-parsexlsx", license: "MIT", committable: true },
  { re: /^xlrd_/, origin: "xlrd", license: "BSD-3-Clause", committable: true },
  { re: /^excel-reader-xlsx_|^jmcnamara_excel-reader-xlsx\//, origin: "excel-reader-xlsx", license: "Artistic-1.0-Perl", committable: false },
  { re: /^pyExcelerator_|^WoLpH_pyExcelerator\//, origin: "pyExcelerator", license: "BSD-4-Clause", committable: false },
  { re: /^jxls-[a-z]+_|^SheetJS_jxls\//, origin: "jxls", license: "LGPL", committable: false },
  { re: /^libreoffice_/, origin: "libreoffice", license: "LGPL-3.0 (per test_files README; upstream test-files are MPL/LGPL)", committable: false },
  { re: /^phpexcel_/, origin: "phpexcel", license: "LGPL-2.1", committable: false },
  { re: /^ootest_/, origin: "openoffice-testdocs", license: "unclear", committable: false },
]

// Stems of SheetJS-original test files described in the test_files README.
const SHEETJS_STEMS = [
  "formula_stress_test", "comments_stress_test", "number_format", "large_strings", "RkNumber", "merge_cells",
  "LONumbers", "rich_text_stress", "named_ranges", "custom_properties", "xlsx-stream-d-date-cell", "pivot_table",
  "AutoFilter", "BlankSheetTypes", "ErrorTypes", "NumberFormatCondition", "protect_stress_test", "calendar_stress_test",
  "hyperlink_stress_test", "internal_link", "hyperlink_no_rels", "text_and_numbers", "cell_style_simple",
  "formulae_test_simple", "password", "sushi", "defined_names", "smart_tags", "A4X", "numfmt", "write",
  "phonetic_text", "row_height", "column_width", "sized_boxen", "sheet_visibility", "page_margins", "outline",
  "author_snowman", "wtf_path", "freeze", "types", "metadata_2019", "minimal_112", "threaded_comments",
  "LO7322", "Attendance", "duration_112",
]
const UNCLEAR = { origin: "unknown", license: "unclear", committable: false }
const SHEETJS = { origin: "sheetjs", license: "Apache-2.0", committable: true }

function provenance(rel) {
  // External origin: check the full relative path and every component's basename-prefix.
  const parts = rel.split("/")
  for (let i = 0; i < parts.length; i++) {
    const tail = parts.slice(i).join("/")
    for (const o of ORIGINS) if (o.re.test(tail)) return { origin: o.origin, license: o.license, committable: o.committable }
  }
  // SheetJS-original only at the root or in the Excel re-save dirs (201*/, biff*/).
  if (parts.length > 2 || (parts.length === 2 && !/^(201\d|biff\d*)$/.test(parts[0]))) return UNCLEAR
  const base = parts[parts.length - 1]
  // stem followed by end, ".", "_", "-" or digits (e.g. a numbered variant like `<stem>5.xls`)
  if (SHEETJS_STEMS.some((st) => base === st || (base.startsWith(st) && /^[._\-\d]/.test(base.slice(st.length))))) return SHEETJS
  if (/^[^_]+$/.test(base)) return SHEETJS // no prefix at all (README: root files are Apache-2.0)
  return UNCLEAR // unknown `<prefix>_`
}

function walk(dir, base = dir, out = []) {
  if (!fs.existsSync(dir)) return out
  for (const d of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, d.name)
    if (d.isDirectory()) walk(p, base, out)
    else if (d.isFile()) out.push(path.relative(base, p).split(path.sep).join("/"))
  }
  return out.sort()
}

// SheetJS's own curated list (tests.lst) — files its test suite reads.
const tested = new Set()
try {
  for (const l of fs.readFileSync(path.join(SHEETJS_DIR, "tests.lst"), "utf8").split("\n")) {
    const s = l.trim()
    if (s && !s.startsWith("#")) tested.add(s)
  }
} catch {}

const entries = []
const others = {}
const add = (source, root, rel) => {
  if (source === "sheetjs" && NON_DATA.test(rel)) return
  if (source !== "sheetjs" && /(^|\/)(README|\.|manifest\.json$)/.test(rel)) return
  const buf = fs.readFileSync(path.join(root, rel))
  let s
  try { s = sniff(buf) } catch (e) { s = { format: "binary", features: { sniffError: String(e?.message ?? e) } } }
  if (!FOCUS.has(s.format) && !allFormats && source === "sheetjs") { others[s.format] = (others[s.format] ?? 0) + 1; return }
  const e = { source, path: rel, format: s.format, size: buf.length, sha256: sha256(buf) }
  if (source === "sheetjs") {
    Object.assign(e, provenance(rel))
    if (tested.has(rel)) e.sheetjsTests = true
    // Excel/SheetJS-derived exports sitting next to their workbook: `<book>.<ext>.<n>.csv`, `<book>.<ext>.xml`
    if (/\.[A-Za-z0-9]+\.(\d+\.(csv|json|rawjson)|xml)$/.test(rel)) e.derived = true
  } else if (source === "hidden-rows") {
    e.origin = "effect-spreadsheet"
    e.license = "MIT"
    e.committable = true
    if (rel.startsWith("renamed/")) e.extensionMismatch = true
  } else {
    e.origin = "effect-spreadsheet"
    e.license = "MIT"
    e.committable = true
  }
  const f = Object.fromEntries(Object.entries(s.features).filter(([, v]) => v !== false && v !== 0 && v != null))
  if (Object.keys(f).length) e.features = f
  entries.push(e)
}
const sj = walk(SHEETJS_ROOT)
if (!sj.length) console.warn("no SheetJS corpus in test/corpus/.cache/sheetjs; run scripts/corpus/fetch.mjs first")
for (const rel of sj) add("sheetjs", SHEETJS_ROOT, rel)
for (const rel of walk(GENERATED_DIR)) add("generated", GENERATED_DIR, rel)
for (const rel of walk(HIDDEN_ROWS_DIR)) add("hidden-rows", HIDDEN_ROWS_DIR, rel)

// Summary
const counts = {}
for (const e of entries) {
  const k = e.format === "xls" ? `xls/biff${e.features?.biff ?? "?"}` : e.format
  const c = (counts[k] ??= { files: 0, unique: new Set(), bytes: 0, hiddenRows: 0, encrypted: 0, generated: 0, "hidden-rows": 0 })
  c.files++
  c.unique.add(e.sha256)
  c.bytes += e.size
  if (e.features?.hiddenRows) c.hiddenRows++
  if (e.features?.encrypted || e.format === "ooxml-encrypted") c.encrypted++
  if (e.source !== "sheetjs") c[e.source] = (c[e.source] ?? 0) + 1
}
const summary = Object.fromEntries(Object.entries(counts).sort().map(([k, c]) => [k, { ...c, unique: c.unique.size }]))

let fetched = null
try { fetched = JSON.parse(fs.readFileSync(path.join(SHEETJS_CACHE, ".fetched.json"), "utf8")) } catch {}
const manifest = {
  version: 1,
  description: "Corpus files for effect-spreadsheet. Paths are relative to the source root: sheetjs => test/corpus/.cache/sheetjs/test_files, generated => test/corpus/generated, hidden-rows => test/corpus/hidden-rows. Format is sniffed from bytes. See docs/corpus.md.",
  sources: {
    sheetjs: { root: "test/corpus/.cache/sheetjs/test_files", url: fetched?.url ?? null, archiveSha256: fetched?.sha256 ?? null, committed: false },
    generated: { root: "test/corpus/generated", committed: true },
    "hidden-rows": { root: "test/corpus/hidden-rows", committed: true, manifest: "test/corpus/hidden-rows/manifest.json" },
  },
  summary,
  excludedFormats: others,
  files: entries,
}
// One file per line keeps diffs readable.
const json = JSON.stringify({ ...manifest, files: [] }, null, 2).replace(/"files": \[\]/, `"files": [\n${entries.map((e) => "    " + JSON.stringify(e)).join(",\n")}\n  ]`)
fs.mkdirSync(CACHE_DIR, { recursive: true }) // a fresh clone has no .cache yet
fs.writeFileSync(MANIFEST_PATH, json + "\n")
console.log(`manifest: ${entries.length} files in ${((Date.now() - t0) / 1000).toFixed(1)}s -> ${path.relative(process.cwd(), MANIFEST_PATH)}`)
console.table(summary)
console.log("excluded (non-focus) formats:", others)

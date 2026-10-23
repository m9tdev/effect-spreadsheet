// Shared helpers for the corpus scripts (Node >= 24, no dependencies).
// Test/script code only: never imported from src/.
//
// - a minimal ZIP reader (central directory + inflateRawSync)
// - a minimal CFB (OLE2) reader
// - format sniffing by magic bytes (never by extension) and cheap feature detection
//
// These readers are deliberately small and only need to be good enough to
// classify corpus files. Every loop is bounded by the buffer length or a counter.

import { createHash } from "node:crypto"
import fs from "node:fs"
import os from "node:os"
import { inflateRawSync } from "node:zlib"
import path from "node:path"
import { fileURLToPath } from "node:url"

export const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..")
export const CORPUS_DIR = path.join(REPO_ROOT, "test", "corpus")
// CORPUS_CACHE_DIR overrides the cache root (used by oracle-selftest.mjs to stay out of the real cache).
export const CACHE_DIR = process.env.CORPUS_CACHE_DIR ? path.resolve(process.env.CORPUS_CACHE_DIR) : path.join(CORPUS_DIR, ".cache")
export const SHEETJS_CACHE = path.join(CACHE_DIR, "sheetjs")
export const GENERATED_DIR = path.join(CORPUS_DIR, "generated")
export const HIDDEN_ROWS_DIR = path.join(CORPUS_DIR, "hidden-rows")
// Full manifest (lists third-party SheetJS corpus files): gitignored cache only.
export const MANIFEST_PATH = path.join(CACHE_DIR, "manifest.json")
// Committed pin of the external corpus (URL + archive sha256), no per-file details.
export const SOURCES_PATH = path.join(CORPUS_DIR, "sources.json")
export const EXPECTED_DIR = path.join(CACHE_DIR, "expected")
// Shared convention with the test oracle: SHEETJS_DIR is the SheetJS repo dir,
// default ~/sheetjs. Each side picks its own entry file in it (here: SHEETJS_ENTRY, xlsx.js).
export const SHEETJS_DIR = process.env.SHEETJS_DIR || path.join(os.homedir(), "sheetjs")
export const SHEETJS_ENTRY = path.join(SHEETJS_DIR, "xlsx.js")

export const sha256 = (buf) => createHash("sha256").update(buf).digest("hex")

// ---------------------------------------------------------------- ZIP

/** List the central directory of a zip. Returns null if not a readable zip. */
export function zipEntries(buf) {
  // EOCD is in the last 22 + 65535 bytes.
  const min = Math.max(0, buf.length - 22 - 0xffff)
  let eocd = -1
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf[i] === 0x50 && buf[i + 1] === 0x4b && buf[i + 2] === 0x05 && buf[i + 3] === 0x06) { eocd = i; break }
  }
  if (eocd < 0) return null
  let count = buf.readUInt16LE(eocd + 10)
  let cdSize = buf.readUInt32LE(eocd + 12)
  let cdOff = buf.readUInt32LE(eocd + 16)
  if (cdOff === 0xffffffff || count === 0xffff) {
    // zip64 locator just before EOCD
    const loc = eocd - 20
    if (loc >= 0 && buf.readUInt32LE(loc) === 0x07064b50) {
      const z64 = Number(buf.readBigUInt64LE(loc + 8))
      if (z64 + 56 <= buf.length && buf.readUInt32LE(z64) === 0x06064b50) {
        count = Number(buf.readBigUInt64LE(z64 + 32))
        cdSize = Number(buf.readBigUInt64LE(z64 + 40))
        cdOff = Number(buf.readBigUInt64LE(z64 + 48))
      }
    }
  }
  if (cdOff + cdSize > buf.length) return null
  const out = []
  let p = cdOff
  // progress: p strictly increases by >= 46 per entry; bounded by count and buffer.
  for (let n = 0; n < count && p + 46 <= buf.length; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break
    const method = buf.readUInt16LE(p + 10)
    let csize = buf.readUInt32LE(p + 20)
    let usize = buf.readUInt32LE(p + 24)
    const nlen = buf.readUInt16LE(p + 28)
    const xlen = buf.readUInt16LE(p + 30)
    const clen = buf.readUInt16LE(p + 32)
    let lho = buf.readUInt32LE(p + 42)
    const name = buf.toString("utf8", p + 46, p + 46 + nlen)
    // zip64 extra field
    let x = p + 46 + nlen
    const xend = x + xlen
    while (x + 4 <= xend) {
      const id = buf.readUInt16LE(x), sz = buf.readUInt16LE(x + 2)
      if (id === 0x0001) {
        let q = x + 4
        if (usize === 0xffffffff && q + 8 <= x + 4 + sz) { usize = Number(buf.readBigUInt64LE(q)); q += 8 }
        if (csize === 0xffffffff && q + 8 <= x + 4 + sz) { csize = Number(buf.readBigUInt64LE(q)); q += 8 }
        if (lho === 0xffffffff && q + 8 <= x + 4 + sz) { lho = Number(buf.readBigUInt64LE(q)); q += 8 }
      }
      x += 4 + sz
    }
    out.push({ name, method, csize, usize, lho })
    p += 46 + nlen + xlen + clen
  }
  return out
}

/** Extract one entry's bytes (stored or deflated). Throws on bad data. */
export function zipRead(buf, e, maxBytes = 512 * 1024 * 1024) {
  const p = e.lho
  if (p + 30 > buf.length || buf.readUInt32LE(p) !== 0x04034b50) throw new Error(`bad local header for ${e.name}`)
  const start = p + 30 + buf.readUInt16LE(p + 26) + buf.readUInt16LE(p + 28)
  const data = buf.subarray(start, start + e.csize)
  if (e.method === 0) return data
  if (e.method === 8) return inflateRawSync(data, { maxOutputLength: maxBytes })
  throw new Error(`unsupported zip method ${e.method} for ${e.name}`)
}

// ---------------------------------------------------------------- CFB

/** Parse a compound file. Returns { entries: [{name, type, start, size}], read(entry) } or null. */
export function cfbOpen(buf) {
  if (buf.length < 512 || buf.readUInt32LE(0) !== 0xe011cfd0 || buf.readUInt32LE(4) !== 0xe11ab1a1) return null
  const ssz = 1 << buf.readUInt16LE(30)
  const mssz = 1 << buf.readUInt16LE(32)
  if (ssz !== 512 && ssz !== 4096) return null
  const nFat = buf.readUInt32LE(44)
  const dirStart = buf.readUInt32LE(48)
  const miniCutoff = buf.readUInt32LE(56)
  const miniFatStart = buf.readUInt32LE(60)
  let difStart = buf.readUInt32LE(68)
  const nDif = buf.readUInt32LE(72)
  const nSectors = Math.floor((buf.length - ssz) / ssz) + 1
  const sectorOff = (s) => ssz + s * ssz
  // FAT sector list from header DIFAT (109) + DIFAT chain
  const fatSecs = []
  for (let i = 0; i < 109 && fatSecs.length < nFat; i++) fatSecs.push(buf.readUInt32LE(76 + i * 4))
  for (let d = 0; d < nDif && difStart < nSectors && fatSecs.length < nFat; d++) {
    const o = sectorOff(difStart)
    const per = ssz / 4 - 1
    for (let i = 0; i < per && fatSecs.length < nFat; i++) fatSecs.push(buf.readUInt32LE(o + i * 4))
    difStart = buf.readUInt32LE(o + per * 4)
  }
  const fat = []
  for (const s of fatSecs) {
    if (s >= nSectors) continue
    const o = sectorOff(s)
    for (let i = 0; i < ssz / 4 && o + i * 4 + 4 <= buf.length; i++) fat.push(buf.readUInt32LE(o + i * 4))
  }
  const chain = (start, table, limit) => {
    const out = []
    const seen = new Set()
    // progress: each step adds a new (unseen) sector; bounded by table length.
    for (let s = start; s < table.length && !seen.has(s) && out.length < limit; s = table[s]) { seen.add(s); out.push(s) }
    return out
  }
  const readChain = (start) => {
    const secs = chain(start, fat, nSectors)
    const parts = secs.filter((s) => sectorOff(s) < buf.length).map((s) => buf.subarray(sectorOff(s), sectorOff(s) + ssz))
    return Buffer.concat(parts)
  }
  const dir = readChain(dirStart)
  const entries = []
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const nl = dir.readUInt16LE(o + 64)
    const type = dir[o + 66]
    if (type === 0) { entries.push(null); continue }
    const name = dir.toString("utf16le", o, o + Math.max(0, Math.min(64, nl) - 2))
    entries.push({ name, type, start: dir.readUInt32LE(o + 116), size: dir.readUInt32LE(o + 120) })
  }
  const root = entries[0]
  let miniStream = null, miniFat = null
  const read = (e) => {
    if (!e) return null
    if (e.size < miniCutoff && e.type === 2) {
      if (!miniStream) {
        miniStream = root ? readChain(root.start) : Buffer.alloc(0)
        const mf = readChain(miniFatStart)
        miniFat = []
        for (let i = 0; i + 4 <= mf.length; i += 4) miniFat.push(mf.readUInt32LE(i))
      }
      const secs = chain(e.start, miniFat, miniFat.length)
      return Buffer.concat(secs.map((s) => miniStream.subarray(s * mssz, s * mssz + mssz))).subarray(0, e.size)
    }
    return readChain(e.start).subarray(0, e.size)
  }
  return { entries: entries.filter(Boolean), read }
}

// ---------------------------------------------------------------- BIFF

const BOF_IDS = new Set([0x0009, 0x0209, 0x0409, 0x0809])

/** BIFF version from a BOF record at offset 0, or null. */
export function biffVersion(buf) {
  if (buf.length < 8) return null
  const id = buf.readUInt16LE(0), len = buf.readUInt16LE(2)
  if (!BOF_IDS.has(id) || len < 4 || len > 20) return null
  if (id === 0x0009) return 2
  if (id === 0x0209) return 3
  if (id === 0x0409) return 4
  const v = buf.readUInt16LE(4)
  if (v === 0x0600) return 8
  if (v === 0x0500) return 5
  if (v === 0x0400) return 4
  if (v === 0x0300) return 3
  if (v === 0x0200 || v === 0x0007) return 2
  // BIFF8 files sometimes carry odd versions; len 16 => BIFF8, 8 => BIFF5
  return len >= 16 ? 8 : 5
}

/** Scan BIFF records for cheap features. Bounded: offset advances by 4 + len each step. */
export function biffFeatures(buf, biff) {
  const f = { biff, encrypted: false, hiddenRows: 0, sheets: 0, hiddenSheets: 0, veryHiddenSheets: 0 }
  let o = 0
  for (let n = 0; o + 4 <= buf.length && n < 50_000_000; n++) {
    const id = buf.readUInt16LE(o), len = buf.readUInt16LE(o + 2)
    const d = o + 4
    if (d + len > buf.length) { f.truncated = true; break }
    if (id === 0x002f) { f.encrypted = true; break } // FILEPASS: what follows is encrypted, stop
    else if (id === 0x0085 && biff >= 5 && len >= 6) { // BOUNDSHEET
      f.sheets++
      const hs = buf[d + 4] & 3
      if (hs === 1) f.hiddenSheets++
      else if (hs === 2) f.veryHiddenSheets++
    } else if (id === 0x0208 && len >= 16 && biff >= 3) { // ROW
      if (buf[d + 12] & 0x20) f.hiddenRows++
    } else if (id === 0x0200 && !f.hugeDimension) { // DIMENSIONS: last row/col + 1
      // BIFF8: rwMic(4) rwMac(4) colMic(2) colMac(2); BIFF2-5: rwMic(2) rwMac(2) colMic(2) colMac(2)
      const big = biff === 8 && len >= 12
      const rwMac = big ? buf.readUInt32LE(d + 4) : len >= 8 ? buf.readUInt16LE(d + 2) : 0
      const colMac = big ? buf.readUInt16LE(d + 10) : len >= 8 ? buf.readUInt16LE(d + 6) : 0
      if (rwMac >= (biff === 8 ? 65536 : 16384) || colMac >= 256) f.hugeDimension = `rows ${rwMac}, cols ${colMac}`
    }
    o = d + len
  }
  return f
}

// ---------------------------------------------------------------- BIFF12 (xlsb)

function readVar(buf, o, max) {
  let v = 0, i = 0
  for (; i < max && o + i < buf.length; i++) {
    const b = buf[o + i]
    v |= (b & 0x7f) << (7 * i)
    if (!(b & 0x80)) return [v >>> 0, i + 1]
  }
  return [v >>> 0, i]
}

/** Scan an xlsb sheet part: hidden BrtRowHdr count and a huge BrtWsDim (reaching row 1048576 or col XFD). */
export function xlsbSheetFeatures(buf) {
  let o = 0, hidden = 0, hugeDimension = null
  for (let n = 0; o < buf.length && n < 100_000_000; n++) {
    const [type, a] = readVar(buf, o, 2)
    const [size, b] = readVar(buf, o + a, 4)
    if (a === 0 || b === 0) break
    const d = o + a + b
    if (type === 0 && size >= 12 && d + 12 <= buf.length) { if (buf[d + 11] & 0x10) hidden++ }
    else if (type === 0x0094 && size >= 16 && d + 16 <= buf.length && !hugeDimension) { // BrtWsDim
      const rwLast = buf.readUInt32LE(d + 4), colLast = buf.readUInt32LE(d + 12)
      if (rwLast >= 1048575 || colLast >= 16383) hugeDimension = `rows ${rwLast + 1}, cols ${colLast + 1}`
    }
    o = d + size // strictly increases (a >= 1)
  }
  return { hidden, hugeDimension }
}

// ---------------------------------------------------------------- text helpers

function decodeHead(buf, n = 4096) {
  const h = buf.subarray(0, n)
  if (h[0] === 0xff && h[1] === 0xfe) return { enc: "utf-16le", bom: true, text: h.subarray(2).toString("utf16le") }
  if (h[0] === 0xfe && h[1] === 0xff) {
    const sw = Buffer.from(h.subarray(2, h.length - (h.length % 2)))
    sw.swap16()
    return { enc: "utf-16be", bom: true, text: sw.toString("utf16le") }
  }
  if (h[0] === 0xef && h[1] === 0xbb && h[2] === 0xbf) return { enc: "utf-8", bom: true, text: h.subarray(3).toString("utf8") }
  return { enc: null, bom: false, text: h.toString("latin1") }
}

function isUtf8(buf) {
  try { new TextDecoder("utf-8", { fatal: true }).decode(buf); return true } catch { return false }
}

function looksText(buf) {
  const n = Math.min(buf.length, 65536)
  let bad = 0
  for (let i = 0; i < n; i++) {
    const c = buf[i]
    if (c === 0) return false
    if (c < 9 || (c > 13 && c < 32 && c !== 26 && c !== 27)) bad++
  }
  return bad <= n / 100
}

function countMatches(s, re) { let n = 0; for (const _ of s.matchAll(re)) n++; return n }

// ---------------------------------------------------------------- sniffing

/**
 * Sniff a file by its bytes. Returns { format, features }.
 * format ∈ xlsx, xlsm, xltx, xlsb, xls (with features.biff), xls-encrypted-ooxml (CFB EncryptedPackage),
 * xlml (SpreadsheetML 2003), ods, fods, sylk, dif, csv, html, numbers, lotus, quattro, dbf, rtf, zip, cfb, xml, text, binary, empty.
 * Feature detection is best effort and may be skipped for big parts (features.partial).
 */
export function sniff(buf) {
  if (buf.length === 0) return { format: "empty", features: {} }
  // ZIP
  if (buf[0] === 0x50 && buf[1] === 0x4b && (buf[2] === 3 || buf[2] === 5 || buf[2] === 7)) return sniffZip(buf)
  // CFB
  if (buf.length >= 8 && buf.readUInt32LE(0) === 0xe011cfd0 && buf.readUInt32LE(4) === 0xe11ab1a1) return sniffCfb(buf)
  // raw BIFF (BIFF2-5 workbooks without a CFB container)
  const bv = biffVersion(buf)
  if (bv) return { format: "xls", features: { ...biffFeatures(buf, bv), container: "raw" } }
  // Lotus 1-2-3 / Quattro (BOF 0x0000 len 2)
  if (buf.length > 6 && buf[0] === 0 && buf[1] === 0 && buf[2] === 2 && buf[3] === 0) return { format: "lotus", features: {} }
  if (buf.length > 6 && buf[0] === 0 && buf[1] === 0 && buf[2] === 0x1a && buf[3] === 0) return { format: "lotus", features: { variant: "wk3+" } }
  // dBASE: version byte + YYMMDD + record count; header length sanity
  if ([0x02, 0x03, 0x30, 0x31, 0x43, 0x63, 0x83, 0x8b, 0xcb, 0xf5, 0xfb].includes(buf[0]) && buf.length > 32 && buf[2] >= 1 && buf[2] <= 12 && buf[3] >= 1 && buf[3] <= 31) {
    const hl = buf.readUInt16LE(8)
    if (hl > 32 && hl < buf.length && buf[hl - 1] === 0x0d) return { format: "dbf", features: {} }
  }
  const head = decodeHead(buf)
  const t = head.text.replace(/^\s+/, "")
  if (/^\{\\rtf/.test(t)) return { format: "rtf", features: {} }
  if (/^ID;/.test(t)) return { format: "sylk", features: sylkFeatures(buf) }
  if (/^TABLE\r?\n0,1/.test(t)) return { format: "dif", features: {} }
  if (/^<(\?xml|[A-Za-z!])/.test(t)) {
    const full = textOf(buf, head)
    if (/urn:schemas-microsoft-com:office:spreadsheet/.test(full.slice(0, 1 << 16)) && /<(ss:)?Workbook[\s>]/.test(full.slice(0, 1 << 16))) {
      return { format: "xlml", features: { encoding: head.enc ?? "utf-8", ...xlmlFeatures(full) } }
    }
    if (/<office:document[\s>]/.test(full.slice(0, 1 << 16)) && /opendocument\.spreadsheet/.test(full.slice(0, 1 << 16))) {
      return { format: "fods", features: odsFeatures(full) }
    }
    if (/^<(!DOCTYPE\s+html|html|table|meta|head|body)/i.test(t)) return { format: "html", features: {} }
    return { format: "xml", features: {} } // some other XML document (not a workbook we know)
  }
  if (/^<(!DOCTYPE\s+html|html|table|meta|head|body)/i.test(t)) return { format: "html", features: {} }
  if (/^[[{]/.test(t) && buf.length < 256 * 1024 * 1024) {
    try { JSON.parse(textOf(buf, head)); return { format: "json", features: {} } } catch {}
  }
  // delimited text
  const textual = head.enc?.startsWith("utf-16") ? true : looksText(buf)
  if (textual) return { format: "csv", features: csvFeatures(buf, head) }
  return { format: "binary", features: {} }
}

function textOf(buf, head) {
  if (head.enc === "utf-16le") return buf.subarray(2).toString("utf16le")
  if (head.enc === "utf-16be") { const b = Buffer.from(buf.subarray(2, buf.length - (buf.length % 2))); b.swap16(); return b.toString("utf16le") }
  return buf.toString(head.bom ? "utf8" : "latin1", head.bom ? 3 : 0)
}

function csvFeatures(buf, head) {
  const text = textOf(buf, head)
  const sample = text.slice(0, 65536)
  const lines = sample.split(/\r\n|\n|\r/).slice(0, 50)
  const counts = { ",": 0, ";": 0, "\t": 0, "|": 0 }
  for (const l of lines) for (const k of Object.keys(counts)) counts[k] += l.split(k).length - 1
  const best = Object.entries(counts).sort((a, b) => b[1] - a[1])[0]
  const f = {
    encoding: head.enc ?? (isUtf8(buf.subarray(0, Math.min(buf.length, 1 << 20))) ? "utf-8" : "8bit"),
    bom: head.bom,
    delimiter: best[1] > 0 ? (best[0] === "\t" ? "tab" : best[0]) : null,
    lineEnding: /\r\n/.test(sample) ? "crlf" : /\r/.test(sample) ? "cr" : "lf",
  }
  // quoted newline: a quote-balanced scan of the sample
  let inQ = false, qnl = false
  for (let i = 0; i < sample.length && !qnl; i++) {
    const c = sample[i]
    if (c === '"') inQ = !inQ
    else if (inQ && (c === "\n" || c === "\r")) qnl = true
  }
  if (qnl) f.quotedNewline = true
  return f
}

function sylkFeatures(buf) {
  const s = buf.toString("latin1")
  // hidden row: F record with M0 (height 0) and R<n> but no C<n> (column)
  let hiddenRows = 0, hiddenCols = 0
  for (const m of s.matchAll(/^F;([^\r\n]*)$/gm)) {
    const fields = m[1].split(";")
    if (!fields.includes("M0")) continue
    if (fields.some((x) => /^R\d+$/.test(x))) hiddenRows++
    else if (fields.some((x) => /^C\d+$/.test(x))) hiddenCols++
  }
  const dim = /^B;([^\r\n]*)$/m.exec(s)
  const by = dim && /(?:^|;)Y(\d+)/.exec(dim[1]), bx = dim && /(?:^|;)X(\d+)/.exec(dim[1])
  const f = {}
  if (hiddenRows) f.hiddenRows = hiddenRows
  if (hiddenCols) f.hiddenCols = hiddenCols
  if ((by && +by[1] >= 1048576) || (bx && +bx[1] >= 16384)) f.hugeDimension = `rows ${by?.[1] ?? "?"}, cols ${bx?.[1] ?? "?"}`
  return f
}

function xlmlFeatures(s) {
  const f = {}
  const sheets = [...s.matchAll(/<(?:ss:)?Worksheet\b([^>]*)>/g)]
  f.sheets = sheets.length
  const hiddenSheets = sheets.filter((m) => /ss:Visible="SheetHidden"|Visible="SheetHidden"/.test(m[1])).length
  const vh = sheets.filter((m) => /Visible="SheetVeryHidden"/.test(m[1])).length
  if (hiddenSheets) f.hiddenSheets = hiddenSheets
  if (vh) f.veryHiddenSheets = vh
  const hr = countMatches(s, /<(?:ss:)?Row\b[^>]*\bss:Hidden="(?:1|true)"/g)
  if (hr) f.hiddenRows = hr
  const er = /\bss:ExpandedRowCount="(\d+)"/.exec(s), ec = /\bss:ExpandedColumnCount="(\d+)"/.exec(s)
  if ((er && +er[1] >= 1048576) || (ec && +ec[1] >= 16384)) f.hugeDimension = `rows ${er?.[1] ?? "?"}, cols ${ec?.[1] ?? "?"}`
  return f
}

function odsFeatures(s) {
  const f = {}
  const hr = countMatches(s, /<table:table-row\b[^>]*table:visibility="(?:collapse|filter)"/g)
  if (hr) f.hiddenRows = hr
  const ht = countMatches(s, /<table:table\b[^>]*table:display="false"/g)
  if (ht) f.hiddenSheets = ht
  if (/<table:table\b[^>]*table:protected="true"/.test(s)) f.protectedSheet = true
  return f
}

// Max inflated size of one zip part we scan. The declared size (usize) is untrusted, so the
// inflate itself is capped too (zipRead maxBytes); a part over the cap is skipped (partial).
const PART_LIMIT = 64 * 1024 * 1024

function sniffZip(buf) {
  let ents
  try { ents = zipEntries(buf) } catch { ents = null }
  if (!ents) return { format: "zip", features: { unreadable: true } }
  const byName = new Map(ents.map((e) => [e.name, e]))
  const read = (name, limit = PART_LIMIT) => {
    const e = byName.get(name)
    if (!e || e.usize > limit) return null
    try { return zipRead(buf, e, limit + 1) } catch { return null }
  }
  const features = { entries: ents.length }
  // ODS
  const mt = read("mimetype")
  if (mt && /opendocument\.spreadsheet/.test(mt.toString("latin1"))) {
    const content = read("content.xml")
    if (content) Object.assign(features, odsFeatures(content.toString("utf8")))
    else features.partial = true
    return { format: "ods", features }
  }
  if (byName.has("Index/Document.iwa") || ents.some((e) => e.name.startsWith("Index/") && e.name.endsWith(".iwa"))) return { format: "numbers", features }
  const ctBuf = read("[Content_Types].xml")
  const ct = ctBuf ? ctBuf.toString("utf8") : ""
  let format = null
  // The main part is the officeDocument target of _rels/.rels; its content type
  // (Override, else Default by extension) decides. Writers such as SheetJS list
  // Defaults for every type (e.g. bin => sheet.binary), so a bare match on
  // [Content_Types].xml is not enough.
  const relsBuf = read("_rels/.rels")
  const rels = relsBuf ? relsBuf.toString("utf8") : ""
  const relM = [...rels.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)].map((m) => m[1]).find((a) => /Type="[^"]*\/officeDocument"/.test(a))
  const target = relM ? /Target="([^"]+)"/.exec(relM)?.[1]?.replace(/^\//, "") : null
  let mainType = null
  if (target) {
    const ov = [...ct.matchAll(/<(?:\w+:)?Override\b([^>]*)>/g)].map((m) => m[1]).find((a) => (/PartName="\/?([^"]+)"/.exec(a)?.[1] ?? "").toLowerCase() === target.toLowerCase())
    if (ov) mainType = /ContentType="([^"]+)"/.exec(ov)?.[1] ?? null
    else {
      const ext = target.split(".").pop().toLowerCase()
      const df = [...ct.matchAll(/<(?:\w+:)?Default\b([^>]*)>/g)].map((m) => m[1]).find((a) => (/Extension="([^"]+)"/.exec(a)?.[1] ?? "").toLowerCase() === ext)
      if (df) mainType = /ContentType="([^"]+)"/.exec(df)?.[1] ?? null
    }
    features.mainPart = target
  }
  const typeFormat = (t) =>
    /sheet\.binary\.macroEnabled\.main/.test(t) ? "xlsb"
    : /(sheet|template|addin)\.macroEnabled\.main\+xml/.test(t) ? "xlsm"
    : /spreadsheetml\.template\.main\+xml/.test(t) ? "xltx"
    : /spreadsheetml\.sheet\.main\+xml/.test(t) ? "xlsx"
    : null
  if (mainType) format = typeFormat(mainType)
  if (!format && target) format = /\.bin$/i.test(target) ? "xlsb" : /workbook\.xml$/i.test(target) ? "xlsx" : null
  if (!format) {
    if (ents.some((e) => /(^|\/)workbook\.xml$/i.test(e.name))) format = "xlsx"
    else if (ents.some((e) => /(^|\/)workbook\.bin$/i.test(e.name))) format = "xlsb"
  }
  if (!format) {
    if (/wordprocessingml|presentationml/.test(ct)) return { format: "ooxml-other", features }
    return { format: "zip", features }
  }
  // Sheet parts and workbook visibility
  try {
    const wbEnt = ents.find((e) => /(^|\/)workbook\.(xml|bin)$/i.test(e.name))
    if (format === "xlsb") {
      let hidden = 0, partial = false
      for (const e of ents) {
        if (!/sheets?\/sheet\d*\.bin$/i.test(e.name)) continue
        if (e.usize > PART_LIMIT) { partial = true; continue }
        let part
        try { part = zipRead(buf, e, PART_LIMIT) } catch { partial = true; continue } // inflates past the cap
        const sf = xlsbSheetFeatures(part)
        hidden += sf.hidden
        if (sf.hugeDimension) features.hugeDimension ??= sf.hugeDimension
      }
      if (hidden) features.hiddenRows = hidden
      if (partial) features.partial = true
    } else {
      if (wbEnt && wbEnt.usize <= PART_LIMIT) {
        const wb = zipRead(buf, wbEnt, PART_LIMIT).toString("utf8")
        const sheets = [...wb.matchAll(/<(?:\w+:)?sheet\b([^>]*?)\/?>/g)]
        features.sheets = sheets.length
        const h = sheets.filter((m) => /\bstate="hidden"/.test(m[1])).length
        const vh = sheets.filter((m) => /\bstate="veryHidden"/.test(m[1])).length
        if (h) features.hiddenSheets = h
        if (vh) features.veryHiddenSheets = vh
      }
      let hidden = 0, partial = false, maxDim = null
      for (const e of ents) {
        if (!/worksheets\/[^/]+\.xml$/i.test(e.name)) continue
        if (e.usize > PART_LIMIT) { partial = true; continue }
        let s
        try { s = zipRead(buf, e, PART_LIMIT).toString("utf8") } catch { partial = true; continue } // inflates past the cap
        hidden += countMatches(s, /<(?:\w+:)?row\b[^>]*\bhidden="(?:1|true)"/g)
        const dm = /<(?:\w+:)?dimension\b[^>]*\bref="([^"]+)"/.exec(s)
        if (dm && /XFD|1048576/.test(dm[1])) maxDim = dm[1]
      }
      if (hidden) features.hiddenRows = hidden
      if (maxDim) features.hugeDimension = maxDim
      if (partial) features.partial = true
    }
  } catch (err) {
    features.error = String(err?.message ?? err).slice(0, 200)
  }
  if (byName.has("xl/vbaProject.bin")) features.vba = true
  return { format, features }
}

function sniffCfb(buf) {
  let cfb
  try { cfb = cfbOpen(buf) } catch { cfb = null }
  if (!cfb) return { format: "cfb", features: { unreadable: true } }
  const names = new Map(cfb.entries.map((e) => [e.name.toLowerCase(), e]))
  if (names.has("encryptedpackage") || names.has("encryptioninfo")) return { format: "ooxml-encrypted", features: { encrypted: true } }
  const wb = names.get("workbook") ?? names.get("book")
  if (wb) {
    let s
    try { s = cfb.read(wb) } catch { s = null }
    if (!s) return { format: "xls", features: { unreadable: true } }
    const v = biffVersion(s) ?? (names.has("workbook") ? 8 : 5)
    return { format: "xls", features: { ...biffFeatures(s, v), container: "cfb" } }
  }
  if (names.has("native") || names.has("nativecontent_main")) return { format: "quattro", features: {} }
  if (names.has("worddocument")) return { format: "cfb-doc", features: {} }
  if ([...names.keys()].some((n) => n.startsWith("ole10native") || n.includes("\u0001ole10native"))) return { format: "cfb", features: {} }
  return { format: "cfb", features: { streams: [...names.keys()].slice(0, 8) } }
}

// ---------------------------------------------------------------- dir compare (for --check modes)

function listFiles(dir, base = dir, out = []) {
  if (!fs.existsSync(dir)) return out
  for (const d of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, d.name)
    if (d.isDirectory()) listFiles(p, base, out)
    else out.push(path.relative(base, p).split(path.sep).join("/"))
  }
  return out
}

/**
 * Compare two directory trees both ways by sha256. `ignore` holds relative paths
 * that may exist only in `committed`. Returns { same, missing, extra, differ, files }.
 */
export function compareDirs(generated, committed, ignore = []) {
  const g = new Set(listFiles(generated)), c = new Set(listFiles(committed))
  const skip = new Set(ignore)
  const missing = [...g].filter((f) => !c.has(f)).sort() // generated but not committed
  const extra = [...c].filter((f) => !g.has(f) && !skip.has(f)).sort() // committed but not generated (stale)
  const differ = [...g].filter((f) => c.has(f) && sha256(fs.readFileSync(path.join(generated, f))) !== sha256(fs.readFileSync(path.join(committed, f)))).sort()
  return { same: !missing.length && !extra.length && !differ.length, missing, extra, differ, files: g.size }
}

export function reportCompare(r, label) {
  for (const f of r.differ) console.log(`DIFF    ${f}`)
  for (const f of r.missing) console.log(`MISSING ${f} (generated, not committed)`)
  for (const f of r.extra) console.log(`EXTRA   ${f} (committed, not generated: stale?)`)
  console.log(r.same ? `${label}: ${r.files} files byte-identical, no missing or extra files` : `${label}: ${r.differ.length} differ, ${r.missing.length} missing, ${r.extra.length} extra`)
  return r.same
}

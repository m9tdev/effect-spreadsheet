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
import { crc32, deflateRawSync, inflateRawSync } from "node:zlib"
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
// An empty SHEETJS_DIR counts as unset (`||`, not `??`), so `SHEETJS_DIR= cmd` uses the default.
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
    out.push({ name, method, csize, usize, lho, crc: buf.readUInt32LE(p + 16) })
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

// BOF ("beginning of file") record, the first record of every BIFF substream.
//
// BIFF8: [MS-XLS] 2.4.21 BOF. The record type is 0x0809 and the record is 16 bytes:
//   vers (2 bytes)     MUST be 0x0600 (BIFF8)
//   dt (2 bytes)       substream type: 0x0005 workbook globals, 0x0010 worksheet or
//                      dialog sheet, 0x0020 chart sheet, 0x0040 macro sheet
//   rupBuild, rupYear (2 bytes each), then 8 bytes of flags
// Earlier versions are outside [MS-XLS]. Their BOF records use other record types and are
// shorter: BIFF2 0x0009, BIFF3 0x0209 and BIFF4 0x0409 (vers + dt, 4-6 bytes). BIFF5 uses
// 0x0809 with vers 0x0500 and an 8-byte record.
const BOF_BIFF8_OR_5 = 0x0809
const BOF_OLDER = new Map([[0x0009, 2], [0x0209, 3], [0x0409, 4]])
const VERS_BIFF8 = 0x0600, VERS_BIFF5 = 0x0500
const BOF_BIFF8_SIZE = 16

/** BIFF version (2, 3, 4, 5 or 8) from a BOF record at offset 0, or null if there isn't one. */
export function biffVersion(buf) {
  if (buf.length < 8) return null
  const type = buf.readUInt16LE(0), size = buf.readUInt16LE(2)
  if (size < 4 || size > 20) return null // vers + dt at least; no BOF is longer than 20 bytes
  if (BOF_OLDER.has(type)) return BOF_OLDER.get(type)
  if (type !== BOF_BIFF8_OR_5) return null
  const vers = buf.readUInt16LE(4)
  if (vers === VERS_BIFF8) return 8
  if (vers === VERS_BIFF5) return 5
  // Any other vers: go by the record size, 16 bytes in BIFF8 and 8 in BIFF5.
  return size >= BOF_BIFF8_SIZE ? 8 : 5
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
const isZip = (b) => b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && b[2] === 3 && b[3] === 4

// Parse a (non-zip64) zip's structure for zipDiff: EOCD, central and local headers, and
// the bytes outside the entries. Returns null when it can't (zip64, data descriptors, ...).
function zipLayout(buf) {
  if (!isZip(buf)) return null
  let eocd = -1
  for (let i = buf.length - 22, min = Math.max(0, buf.length - 22 - 0xffff); i >= min; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) return null
  const count = buf.readUInt16LE(eocd + 10), cdSize = buf.readUInt32LE(eocd + 12), cdOff = buf.readUInt32LE(eocd + 16)
  const commentLen = buf.readUInt16LE(eocd + 20)
  if (count === 0xffff || cdOff === 0xffffffff || cdOff + cdSize !== eocd) return null
  const entries = []
  let p = cdOff
  // progress: p increases by >= 46 per entry, bounded by count
  for (let n = 0; n < count; n++) {
    if (p + 46 > eocd || buf.readUInt32LE(p) !== 0x02014b50) return null
    const nlen = buf.readUInt16LE(p + 28), xlen = buf.readUInt16LE(p + 30), clen = buf.readUInt16LE(p + 32)
    const lho = buf.readUInt32LE(p + 42)
    if (lho + 30 > buf.length || buf.readUInt32LE(lho) !== 0x04034b50) return null
    // within one zip, the local header's method and compressed size must agree with the central one
    if (buf.readUInt16LE(lho + 8) !== buf.readUInt16LE(p + 10) || buf.readUInt32LE(lho + 18) !== buf.readUInt32LE(p + 20)) return null
    const lnlen = buf.readUInt16LE(lho + 26), lxlen = buf.readUInt16LE(lho + 28)
    const flags = buf.readUInt16LE(p + 8)
    if (flags & 0x0008) return null // data descriptor: not used by our writer
    entries.push({
      // everything in the central header except method (10), csize (20) and offset (42)
      central: Buffer.concat([buf.subarray(p + 4, p + 10), buf.subarray(p + 12, p + 20), buf.subarray(p + 24, p + 42), buf.subarray(p + 46, p + 46 + nlen + xlen + clen)]),
      // everything in the local header except method (8) and csize (18)
      local: Buffer.concat([buf.subarray(lho + 4, lho + 8), buf.subarray(lho + 10, lho + 18), buf.subarray(lho + 22, lho + 30 + lnlen + lxlen)]),
      name: buf.toString("utf8", p + 46, p + 46 + nlen),
      crcCentral: buf.readUInt32LE(p + 16), crcLocal: buf.readUInt32LE(lho + 14),
      usizeCentral: buf.readUInt32LE(p + 24), usizeLocal: buf.readUInt32LE(lho + 22),
      e: { name: buf.toString("utf8", p + 46, p + 46 + nlen), method: buf.readUInt16LE(p + 10), csize: buf.readUInt32LE(p + 20), usize: buf.readUInt32LE(p + 24), lho },
      dataEnd: lho + 30 + lnlen + lxlen + buf.readUInt32LE(p + 20),
      lho,
    })
    p += 46 + nlen + xlen + clen
  }
  if (p !== eocd) return null // junk at the end of the central directory
  // local entries must sit back to back from offset 0 up to the central directory, in central order
  let at = 0
  for (const x of entries) { if (x.lho !== at) return null; at = x.dataEnd }
  if (at !== cdOff) return null
  return {
    entries,
    eocd: Buffer.concat([buf.subarray(eocd + 4, eocd + 12)]), // disk numbers and entry counts
    comment: buf.subarray(eocd + 22, eocd + 22 + commentLen),
    trailing: buf.subarray(Math.min(buf.length, eocd + 22 + commentLen)),
    commentLenOk: eocd + 22 + commentLen <= buf.length,
  }
}

/**
 * Why two zips differ in more than their compressed bytes, or null when they don't.
 * Deflate output isn't stable across zlib builds (Node's bundled zlib vs a system zlib,
 * CPU-specific code paths), so a regenerated zip may legitimately differ in each entry's
 * compression method, compressed size, compressed data and hence its offsets. Everything
 * else must be equal: every other local and central header field (timestamps, CRC32, sizes,
 * flags, attributes, names, extra fields, comments), the entry order, the archive comment
 * and any trailing bytes. Each entry's CRC32 and uncompressed size are also checked against
 * its uncompressed bytes.
 */
export function zipDiff(a, b) {
  const la = zipLayout(a), lb = zipLayout(b)
  if (!la || !lb) return "not a plain zip (zip64, data descriptor or bad layout)"
  if (la.entries.length !== lb.entries.length) return `entry count ${la.entries.length} vs ${lb.entries.length}`
  if (!la.eocd.equals(lb.eocd)) return "end-of-central-directory fields"
  if (!la.commentLenOk || !lb.commentLenOk) return "archive comment runs past the end"
  if (!la.comment.equals(lb.comment)) return "archive comment"
  if (!la.trailing.equals(lb.trailing)) return "trailing bytes after the archive"
  for (let i = 0; i < la.entries.length; i++) {
    const x = la.entries[i], y = lb.entries[i]
    if (x.name !== y.name) return `entry ${i}: name ${x.name} vs ${y.name}`
    if (!x.central.equals(y.central)) return `${x.name}: central header fields`
    if (!x.local.equals(y.local)) return `${x.name}: local header fields`
    for (const [buf, z, label] of [[a, x, "first"], [b, y, "second"]]) {
      let data
      try { data = Buffer.from(zipRead(buf, z.e)) } catch (err) { return `${z.name}: unreadable in the ${label} zip (${err.message})` }
      const c = crc32(data)
      if (c !== z.crcCentral || c !== z.crcLocal) return `${z.name}: CRC32 doesn't match the data in the ${label} zip`
      if (data.length !== z.usizeCentral || data.length !== z.usizeLocal) return `${z.name}: uncompressed size doesn't match the data in the ${label} zip`
      z.data = data
    }
    if (!x.data.equals(y.data)) return `${x.name}: uncompressed bytes`
  }
  return null
}

/** Same zip apart from compressed bytes (see zipDiff). */
export const sameZipContent = (a, b) => zipDiff(a, b) === null

/**
 * Compare two directory trees both ways. `ignore` holds relative paths that may exist only
 * in `committed`. Files match when their bytes are equal, or (zips) when their content is
 * (sameZipContent), or when `equal(rel, generatedBuf, committedBuf)` says so.
 * Returns { same, missing, extra, differ, zipRecompressed, files }.
 */
export function compareDirs(generated, committed, ignore = [], equal = null) {
  const g = new Set(listFiles(generated)), c = new Set(listFiles(committed))
  const skip = new Set(ignore)
  const missing = [...g].filter((f) => !c.has(f)).sort() // generated but not committed
  const extra = [...c].filter((f) => !g.has(f) && !skip.has(f)).sort() // committed but not generated (stale)
  const differ = [], zipRecompressed = [], reasons = {}
  for (const f of [...g].filter((f) => c.has(f)).sort()) {
    const a = fs.readFileSync(path.join(generated, f)), b = fs.readFileSync(path.join(committed, f))
    if (a.equals(b)) continue
    const zd = isZip(a) && isZip(b) ? zipDiff(a, b) : "not zips"
    if (zd === null) { zipRecompressed.push(f); continue }
    if (equal?.(f, a, b)) continue
    differ.push(f)
    if (zd !== "not zips") reasons[f] = zd
  }
  return { same: !missing.length && !extra.length && !differ.length, missing, extra, differ, zipRecompressed, reasons, files: g.size }
}

export function reportCompare(r, label) {
  for (const f of r.differ) console.log(`DIFF    ${f}${r.reasons?.[f] ? `  (${r.reasons[f]})` : ""}`)
  for (const f of r.missing) console.log(`MISSING ${f} (generated, not committed)`)
  for (const f of r.extra) console.log(`EXTRA   ${f} (committed, not generated: stale?)`)
  // zipDiff === null: only compression method/size/data (and so offsets) differ, which is what another zlib build does
  const z = r.zipRecompressed?.length ? ` (${r.zipRecompressed.length} zips differ only in their compressed bytes: another zlib build)` : ""
  console.log(r.same ? `${label}: ${r.files} files identical, no missing or extra files${z}` : `${label}: ${r.differ.length} differ, ${r.missing.length} missing, ${r.extra.length} extra${z}`)
  return r.same
}

// ---------------------------------------------------------------- argument checks

/**
 * Parse simple CLI flags. `flags` lists boolean flags, `withValue` flags that take a
 * value. Anything else (e.g. a typo like --chek) exits with code 2.
 */
export function parseFlags(argv, flags, withValue = []) {
  const out = {}
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i]
    if (flags.includes(k)) out[k] = true
    else if (withValue.includes(k)) {
      const v = argv[i + 1]
      if (v === undefined || v.startsWith("--")) { console.error(`${k} needs a value`); process.exit(2) }
      out[k] = v
      i++
    } else {
      console.error(`unknown argument ${k} (allowed: ${[...flags, ...withValue.map((w) => w + " VALUE")].join(", ") || "none"})`)
      process.exit(2)
    }
  }
  return out
}

// ---------------------------------------------------------------- zip writer (fixture generators)

/**
 * Zip with a fixed DOS timestamp (1980-01-01), deflate (level 9 by default) unless that's not
 * smaller. Deterministic for a given zlib build; other zlib builds may deflate differently.
 */
export function zip(files, { level = 9 } = {}) {
  const locals = [], centrals = []
  let off = 0
  for (const [name, content] of files) {
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, "utf8")
    const nameBuf = Buffer.from(name, "utf8")
    const deflated = deflateRawSync(data, { level })
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

/**
 * Re-zip an existing zip deterministically.
 * transform: {
 *   drop?: (name) => boolean            // parts to leave out
 *   parts?: { [name]: (Buffer) => Buffer } // per-part byte transforms
 *   replace?: Array<[name | RegExp, string | RegExp, string]> // string replacements in UTF-8 parts
 * }
 */
export function rezip(buf, transform = {}) {
  const { drop = () => false, parts = {}, replace = [] } = transform
  const files = []
  for (const e of zipEntries(buf)) {
    if (e.name.endsWith("/") || drop(e.name)) continue
    let data = Buffer.from(zipRead(buf, e))
    if (parts[e.name]) data = parts[e.name](data)
    const reps = replace.filter(([p]) => (typeof p === "string" ? p === e.name : p.test(e.name)))
    if (reps.length) {
      let s = data.toString("utf8")
      for (const [, from, to] of reps) s = s.replace(from, to)
      data = Buffer.from(s, "utf8")
    }
    files.push([e.name, data])
  }
  return zip(files)
}

// ---------------------------------------------------------------- strip SheetJS-writer boilerplate

function writeVar(v, maxBytes) {
  const out = []
  for (let i = 0; i < maxBytes; i++) {
    const b = v & 0x7f
    v >>>= 7
    if (v === 0) { out.push(b); return Buffer.from(out) }
    out.push(b | 0x80)
  }
  throw new Error("varint too large")
}

/** Replace the BrtFileVersion record (type 0x80, [MS-XLSB] 2.4.649) with empty app/version strings. */
function xlsbBlankFileVersion(buf) {
  const parts = []
  let o = 0
  // progress: o strictly increases by header + size per record
  for (let n = 0; o < buf.length && n < 10_000_000; n++) {
    const [type, a] = readVar(buf, o, 2)
    const [size, b] = readVar(buf, o + a, 4)
    if (a === 0 || b === 0) throw new Error("bad BIFF12 record header")
    const end = o + a + b + size
    if (type === 0x0080) {
      // guidCodeName (16 bytes), then stAppName, stLastEdited, stLowestEdited, stRupBuild (XLWideString: u32 count + UTF-16)
      const data = Buffer.alloc(16 + 4 * 4)
      parts.push(writeVar(type, 2), writeVar(data.length, 4), data)
    } else parts.push(buf.subarray(o, end))
    o = end
  }
  return Buffer.concat(parts)
}

/** Remove the SheetJS writer's own parts/strings from an xlsx/xlsb/ods zip. Parse results don't change. */
export function stripWriterZip(buf) {
  return rezip(buf, {
    drop: (n) => n === "xl/theme/theme1.xml",
    parts: { "xl/workbook.bin": xlsbBlankFileVersion },
    replace: [
      ["[Content_Types].xml", /<Override PartName="\/xl\/theme\/theme1\.xml"[^>]*\/>/, ""],
      [/^xl\/_rels\/workbook\.(xml|bin)\.rels$/, /<Relationship [^>]*Target="theme\/theme1\.xml"\/>/, ""],
      ["docProps/app.xml", "<Application>SheetJS</Application>", ""],
      ["meta.xml", /<meta:generator>SheetJS[^<]*<\/meta:generator>/, ""],
      ["content.xml", / xmlns:sheet="urn:oasis:names:tc:opendocument:sh33tjs:1\.0"/, ""],
      ["styles.xml", / xmlns:sheet="urn:oasis:names:tc:opendocument:sh33tjs:1\.0"/, ""],
    ],
  })
}

/** Same for text formats: SYLK ID record, flat-ODS generator and namespace. */
export function stripWriterText(s) {
  return s
    .replace(/^ID;PSheetJS;N;E(\r?\n)/, "ID;PWXL;N;E$1")
    .replace(/<meta:generator>SheetJS[^<]*<\/meta:generator>/, "")
    .replace(/ xmlns:sheet="urn:oasis:names:tc:opendocument:sh33tjs:1\.0"/, "")
}

/**
 * Unlink a stream from a compound file's directory ([MS-CFB] 2.6) and zero its data.
 * Only handles an entry with at most one child (true for the SheetJS marker stream).
 * Returns a patched copy, or the input when the entry isn't there.
 */
export function cfbUnlink(input, name) {
  const buf = Buffer.from(input)
  const ssz = 1 << buf.readUInt16LE(30), mssz = 1 << buf.readUInt16LE(32)
  const nFat = buf.readUInt32LE(44), dirStart = buf.readUInt32LE(48), miniCutoff = buf.readUInt32LE(56), miniFatStart = buf.readUInt32LE(60)
  if (buf.readUInt32LE(72) !== 0) throw new Error("cfbUnlink: DIFAT sectors not supported")
  const off = (s) => ssz + s * ssz
  const fat = []
  for (let i = 0; i < Math.min(nFat, 109); i++) {
    const s = buf.readUInt32LE(76 + i * 4)
    for (let j = 0; j < ssz / 4; j++) fat.push(buf.readUInt32LE(off(s) + j * 4))
  }
  const chain = (start, table) => { const out = [], seen = new Set(); for (let s = start; s < table.length && !seen.has(s); s = table[s]) { seen.add(s); out.push(s) } return out }
  const ents = chain(dirStart, fat).flatMap((s) => Array.from({ length: ssz / 128 }, (_, k) => off(s) + k * 128))
  const nameOf = (e) => buf.toString("utf16le", e, e + Math.max(0, Math.min(64, buf.readUInt16LE(e + 64)) - 2))
  const idx = ents.findIndex((e) => buf[e + 66] !== 0 && nameOf(e) === name)
  if (idx < 0) return buf
  const v = ents[idx]
  const L = buf.readInt32LE(v + 68), R = buf.readInt32LE(v + 72), C = buf.readInt32LE(v + 76)
  if (C !== -1 || (L !== -1 && R !== -1)) throw new Error("cfbUnlink: entry has children on both sides")
  const repl = L !== -1 ? L : R
  for (const e of ents) for (const o of [68, 72, 76]) if (buf[e + 66] !== 0 && buf.readInt32LE(e + o) === idx) buf.writeInt32LE(repl, e + o)
  // zero the stream data
  const start = buf.readUInt32LE(v + 116), size = buf.readUInt32LE(v + 120)
  if (size < miniCutoff) {
    const miniFat = chain(miniFatStart, fat).flatMap((s) => Array.from({ length: ssz / 4 }, (_, j) => buf.readUInt32LE(off(s) + j * 4)))
    const root = ents[0], rootSecs = chain(buf.readUInt32LE(root + 116), fat)
    for (const ms of chain(start, miniFat)) {
      const pos = ms * mssz, sec = rootSecs[Math.floor(pos / ssz)]
      if (sec !== undefined) buf.fill(0, off(sec) + (pos % ssz), off(sec) + (pos % ssz) + mssz)
    }
  } else for (const s of chain(start, fat)) buf.fill(0, off(s), off(s) + ssz)
  buf.fill(0, v, v + 128)
  buf.writeInt32LE(-1, v + 68); buf.writeInt32LE(-1, v + 72); buf.writeInt32LE(-1, v + 76)
  return buf
}

/** xls written by SheetJS: drop the \x01Sh33tJ5 marker stream, blank "Sh33tJS" in WRITEACCESS and the "7262" SystemIdentifier. */
export function stripWriterXls(buf) {
  const out = cfbUnlink(buf, "\u0001Sh33tJ5")
  const tag = Buffer.from("Sh33tJS", "latin1")
  // progress: each step moves i past the previous match
  for (let i = out.indexOf(tag), n = 0; i >= 0 && n < 4; i = out.indexOf(tag, i + tag.length), n++) out.fill(0x20, i, i + tag.length)
  // Property-set streams ([MS-OLEPS] 2.21 PropertySetStream): ByteOrder FFFE, Version 0,
  // then SystemIdentifier, which SheetJS fills with "7262". Zero it.
  const ps = Buffer.from([0xfe, 0xff, 0x00, 0x00, 0x37, 0x32, 0x36, 0x32])
  for (let i = out.indexOf(ps), n = 0; i >= 0 && n < 4; i = out.indexOf(ps, i + ps.length), n++) out.fill(0, i + 4, i + 8)
  return out
}

/** Strip SheetJS-writer boilerplate from any fixture, picked by its bytes (no-op for others). */
export function stripWriterBoilerplate(buf) {
  if (buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 3) return stripWriterZip(buf)
  if (buf.length >= 512 && buf.readUInt32LE(0) === 0xe011cfd0) return stripWriterXls(buf)
  const head = buf.subarray(0, 8).toString("latin1")
  if (head.startsWith("ID;") || head.startsWith("<?xml")) {
    const enc = head.startsWith("ID;") ? "latin1" : "utf8"
    const s = buf.toString(enc), t = stripWriterText(s)
    return t === s ? buf : Buffer.from(t, enc)
  }
  // raw BIFF2-4: WRITEACCESS-like username strings
  const tag = Buffer.from("Sh33tJS", "latin1")
  if (buf.indexOf(tag) >= 0) { const b = Buffer.from(buf); for (let i = b.indexOf(tag); i >= 0; i = b.indexOf(tag, i + 7)) b.fill(0x20, i, i + 7); return b }
  return buf
}

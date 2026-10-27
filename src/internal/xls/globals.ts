/**
 * The workbook globals substream of a BIFF8 or BIFF5 workbook stream: sheet
 * list (BOUNDSHEET), code page, date system, number formats (FORMAT) and
 * cell styles (XF), the shared-string table (SST with CONTINUE), defined
 * names (NAME with EXTERNSHEET and SUPBOOK) and encryption (FILEPASS).
 *
 * The walk reads records from the first BOF to its EOF; every record
 * advances the reader. Limits: sheets, defined names, shared strings, string lengths, and
 * the SST's and each SUPBOOK's total size (maxDecompressedBytes).
 *
 * @internal
 */
import { Encrypted, LimitExceeded, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { BOF_BIFF5_8, type ContinuedRecord, EOF, RecordReader, readBof } from "../biff/records.ts"
import { readRichString, SegmentReader, utf16 } from "../biff/segments.ts"
import { formatFor, type NumberFormat } from "../xlsx/numFmt.ts"
import { codepageLabel, Fields, readString, type Text, textDecoder } from "./fields.ts"

const BOUNDSHEET = 0x0085
const CODEPAGE = 0x0042
const DATEMODE = 0x0022
const FORMAT = 0x041E
const XF = 0x00E0
const SST = 0x00FC
const FILEPASS = 0x002F
const NAME = 0x0018
const EXTERNSHEET = 0x0017
const SUPBOOK = 0x01AE

export type SheetKind = "worksheet" | "macrosheet" | "chartsheet" | "other"

export interface SheetEntry {
  readonly name: string
  readonly visibility: "visible" | "hidden" | "veryHidden"
  /** Offset of the sheet's BOF in the workbook stream. */
  readonly offset: number
  /**
   * Where its substream must end: the next substream's offset (in offset
   * order, whatever the BOUNDSHEET order), or the stream's end. Substreams
   * don't overlap, so all sheets together read the stream at most once.
   */
  readonly end: number
  readonly kind: SheetKind
}

export interface DefinedNameEntry {
  readonly name: string
  readonly formula: string
  readonly scope: number | undefined
  readonly hidden: boolean
}

export interface Globals {
  readonly biff8: boolean
  readonly text: Text
  readonly date1904: boolean
  readonly sheets: ReadonlyArray<SheetEntry>
  /** Number format by XF index (a cell's `ixfe`). */
  readonly formats: ReadonlyArray<NumberFormat>
  readonly strings: ReadonlyArray<string>
  readonly definedNames: ReadonlyArray<DefinedNameEntry>
}

/** A string longer than maxStringLength (`actual` when its length is known). */
const tooLong = (limits: Limits, actual = limits.maxStringLength + 1) =>
  new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual })

const malformed = (reason: string, offset?: number) =>
  new MalformedFile({ format: "xls", part: "workbook globals", reason, ...(offset === undefined ? {} : { offset }) })

/** Built-in defined names (NAME with fBuiltin): one character code each (MS-XLS 2.5.164 Lbl). */
const BUILTIN_NAMES = [
  "Consolidate_Area", "Auto_Open", "Auto_Close", "Extract", "Database", "Criteria", "Print_Area",
  "Print_Titles", "Recorder", "Data_Form", "Auto_Activate", "Auto_Deactivate", "Sheet_Title", "_FilterDatabase"
]

/**
 * A SUPBOOK record: the workbook itself, an add-in's functions, or an
 * external workbook (or DDE/OLE link). `link` numbers the external ones
 * from 1, in record order, as XLSX numbers its external links (`[1]`).
 */
type Book =
  | { readonly kind: "self" }
  | { readonly kind: "addin" }
  | { readonly kind: "external"; readonly link: number; readonly sheets: ReadonlyArray<string> }

const SELF: Book = { kind: "self" }

/** An EXTERNSHEET entry (XTI): a SUPBOOK and a range of its sheets. */
interface Xti {
  readonly book: number
  readonly first: number
  readonly last: number
}

/** Reads a SUPBOOK (with its CONTINUEs): `ctab`, then a marker, or a path and `ctab` sheet names. */
const readSupbook = (record: ContinuedRecord, link: number, maxLength: number): Book | null => {
  // this workbook and add-ins are exactly ctab + marker (4 bytes); an
  // external book's path can have a character count equal to a marker
  const marker = record.data.length === 4 ? new Fields(record.data).u16(2) : undefined
  if (marker === 0x0401) return SELF
  if (marker === 0x3A01) return { kind: "addin" }
  const segments = new SegmentReader(record.segments)
  const ctab = segments.u16() ?? 0
  const path = readRichString(segments, maxLength) // the path's character count is the marker's place
  if (path === null) return null
  const sheets: Array<string> = []
  // progress: k up to ctab; each name consumes at least 3 bytes, or the loop ends
  for (let k = 0; path !== undefined && k < ctab; k++) {
    const name = readRichString(segments, maxLength)
    if (name === null) return null
    if (name === undefined) break // fewer names than declared: keep those read
    sheets.push(name)
  }
  return { kind: "external", link, sheets }
}

/**
 * Whether a sheet name needs quotes in a reference, as Excel writes them:
 * when it isn't name-like (a letter or `_` first, then letters, digits, `_`
 * and `.`; characters beyond ASCII count as letters), or when it reads as
 * a cell reference (`A1`, up to column XFD) or an R1C1 one (`R`, `C2`,
 * `R1C1`).
 */
const needsQuotes = (name: string): boolean => {
  if (!/^[A-Za-z_\u0080-\uFFFF][A-Za-z0-9_.\u0080-\uFFFF]*$/.test(name)) return true
  if (/^(?:R\d*|C\d*|R\d*C\d*)$/i.test(name)) return true
  const a1 = /^([A-Za-z]{1,3})\d+$/.exec(name)
  return a1 !== null && columnNumber(a1[1]!) <= 16384
}

/** A sheet prefix (`Sheet!`, `'My Sheet'!`, `Jan:Mar!`, `[1]Data!`): quoted as a whole when any sheet name needs it. */
const sheetPrefix = (book: string, names: ReadonlyArray<string>): string => {
  const text = book + names.join(":")
  return names.some(needsQuotes) ? `'${text.replace(/'/g, "''")}'!` : `${text}!`
}

/** The column number of letters (`A` 1, `XFD` 16384). */
const columnNumber = (letters: string): number => {
  let n = 0
  // progress: k up to letters.length
  for (let k = 0; k < letters.length; k++) n = n * 26 + (letters.toUpperCase().charCodeAt(k) - 64)
  return n
}

const column = (c: number): string => {
  let out = ""
  let n = c + 1
  // progress: n shrinks by a factor of 26
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** A column or row of a BIFF8 reference: its column word has bit 14 for a relative column, bit 15 for a relative row. */
const colText = (col: number): string => `${(col & 0x4000) !== 0 ? "" : "$"}${column(col & 0x3FFF)}`
const rowText = (row: number, col: number): string => `${(col & 0x8000) !== 0 ? "" : "$"}${row + 1}`
const cellText = (row: number, col: number): string => colText(col) + rowText(row, col)

/** BIFF8 sheets: rows 0-65535, columns 0-255. */
const LAST_ROW = 0xFFFF
const LAST_COL = 0xFF

/** An area: whole rows as `$1:$2`, whole columns as `$A:$B`, else `$A$1:$B$2`. */
const areaText = (r1: number, r2: number, c1: number, c2: number): string => {
  if ((c1 & 0x3FFF) === 0 && (c2 & 0x3FFF) === LAST_COL) return `${rowText(r1, c1)}:${rowText(r2, c2)}`
  if (r1 === 0 && r2 === LAST_ROW) return `${colText(c1)}:${colText(c2)}`
  return `${cellText(r1, c1)}:${cellText(r2, c2)}`
}

/**
 * The formula of a defined name, for the forms names usually take: a
 * single 3-D reference or area, written as Excel writes it in XLSX
 * (`Sheet!$A$1`, `'My Sheet'!$A$1:$B$2`, `Jan:Mar!$A$1`, `[1]Data!$A:$A`).
 * A reference to a deleted or missing sheet gives `#REF!$A$1`; one to
 * deleted cells (PtgRefErr3d, PtgAreaErr3d) `Sheet!#REF!`, or `#REF!` when
 * its sheet is gone too; one whose
 * EXTERNSHEET entry is missing gives `""`, like anything else (formulas
 * aren't decompiled). With no SUPBOOK at all (writers that leave it out),
 * the entries of SUPBOOK 0 are taken to point at this workbook.
 */
const nameFormula = (
  rgce: Uint8Array,
  xtis: ReadonlyArray<Xti>,
  books: ReadonlyArray<Book>,
  sheets: ReadonlyArray<SheetEntry>
): string => {
  const f = new Fields(rgce)
  const ptg = f.u8(0)
  const prefix = (ixti: number): string | undefined => {
    const xti = xtis[ixti]
    if (xti === undefined) return undefined
    const book = books.length === 0 && xti.book === 0 ? SELF : books[xti.book]
    if (book === undefined || book.kind === "addin") return "#REF!"
    const names = book.kind === "self" ? sheets.map((s) => s.name) : book.sheets
    const first = names[xti.first]
    const last = names[xti.last]
    if (first === undefined || last === undefined || xti.last < xti.first) return "#REF!"
    return sheetPrefix(book.kind === "self" ? "" : `[${book.link}]`, xti.first === xti.last ? [first] : [first, last])
  }
  if ((ptg === 0x3A || ptg === 0x5A || ptg === 0x7A) && rgce.length === 7) {
    const sheet = prefix(f.u16(1)!)
    return sheet === undefined ? "" : sheet + cellText(f.u16(3)!, f.u16(5)!)
  }
  if ((ptg === 0x3B || ptg === 0x5B || ptg === 0x7B) && rgce.length === 11) {
    const sheet = prefix(f.u16(1)!)
    return sheet === undefined ? "" : sheet + areaText(f.u16(3)!, f.u16(5)!, f.u16(7)!, f.u16(9)!)
  }
  // PtgRefErr3d / PtgAreaErr3d: a reference whose cells were deleted
  if (((ptg === 0x3C || ptg === 0x5C || ptg === 0x7C) && rgce.length === 7) || ((ptg === 0x3D || ptg === 0x5D || ptg === 0x7D) && rgce.length === 11)) {
    const sheet = prefix(f.u16(1)!)
    return sheet === undefined ? "" : sheet === "#REF!" ? sheet : `${sheet}#REF!`
  }
  return ""
}

/**
 * Each sheet's substream end (see `SheetEntry.end`), after checking that
 * no sheet starts inside the globals (which end at `globalsEnd`) and that
 * no two sheets share a substream: either would let a small file make the
 * reader walk the same records once per sheet. VB module offsets count as
 * substream starts only where they point at a BOF past the globals (they
 * are often 0 or stale: a module's code lives outside the workbook stream).
 */
const sheetEnds = (
  stream: Uint8Array,
  globalsEnd: number,
  sheets: ReadonlyArray<Omit<SheetEntry, "end">>,
  modules: ReadonlyArray<number>
): Array<SheetEntry> | MalformedFile => {
  const owner = new Map<number, string>()
  // progress: one sheet per iteration (at most maxSheets)
  for (const sheet of sheets) {
    if (sheet.offset < globalsEnd) {
      return malformed(`sheet ${JSON.stringify(sheet.name)} starts at ${sheet.offset}, inside the workbook globals (which end at ${globalsEnd})`)
    }
    const other = owner.get(sheet.offset)
    if (other !== undefined) {
      return malformed(`sheets ${JSON.stringify(other)} and ${JSON.stringify(sheet.name)} both start at ${sheet.offset}`)
    }
    owner.set(sheet.offset, sheet.name)
  }
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength)
  const unique = new Set(owner.keys())
  // progress: one module per iteration (bounded by the BOUNDSHEET records read)
  for (const offset of modules) {
    if (offset >= globalsEnd && offset + 4 <= stream.length && view.getUint16(offset, true) === BOF_BIFF5_8) unique.add(offset)
  }
  const starts = Float64Array.from(unique).sort()
  const ends = new Map<number, number>()
  // progress: k up to starts.length
  for (let k = 0; k < starts.length; k++) ends.set(starts[k]!, k + 1 < starts.length ? starts[k + 1]! : stream.length)
  return sheets.map((sheet) => ({ ...sheet, end: ends.get(sheet.offset)! }))
}

/** Reads the globals of a workbook stream (its first substream). */
export const readGlobals = (stream: Uint8Array, limits: Limits): Globals | SpreadsheetError => {
  const reader = new RecordReader(stream, "workbook globals")
  const first = reader.next()
  if (first === undefined) return malformed("empty workbook stream")
  if (first instanceof MalformedFile) return first
  const bof = readBof(first)
  if (bof === undefined || bof.kind !== 0x0005) return malformed("the workbook stream doesn't start with a workbook BOF", 0)
  // BIFF5 writers put 0x0500 or another 0x05xx (e.g. 0x0580) here
  if (bof.version !== 0x0600 && (bof.version & 0xFF00) !== 0x0500) {
    return new UnsupportedFormat({ format: "xls", reason: `BIFF version 0x${bof.version.toString(16)} in a BIFF5/8 BOF` })
  }
  const biff8 = bof.version === 0x0600
  let text: Text = { biff8, decoder: textDecoder("windows-1252"), maxLength: limits.maxStringLength }
  let date1904 = false
  const sheets: Array<Omit<SheetEntry, "end">> = []
  /** Offsets of VB module BOUNDSHEETs: not sheets, but they may start a substream. */
  const modules: Array<number> = []
  const custom = new Map<number, string>()
  const xfFormats: Array<number> = []
  let strings: Array<string> = []
  const names: Array<{ name: string; hidden: boolean; scope: number | undefined; rgce: Uint8Array }> = []
  const books: Array<Book> = []
  let externalBooks = 0
  let xtis: Array<Xti> = []

  // progress: every iteration reads one record (the reader's offset grows)
  for (;;) {
    const type = reader.peekType()
    if (type === undefined) return malformed("the workbook globals have no EOF")
    if (type === SST) {
      if (!biff8) {
        reader.next()
        continue
      }
      const sst = reader.nextWithContinues(limits.maxDecompressedBytes, "maxDecompressedBytes")
      if (sst === undefined) return malformed("truncated SST")
      if (sst instanceof MalformedFile || sst instanceof LimitExceeded) return sst
      const segments = new SegmentReader(sst.segments)
      segments.u32() // total count
      const unique = segments.u32()
      if (unique === undefined) return malformed("truncated SST", sst.offset)
      if (unique > limits.maxSharedStrings) {
        return new LimitExceeded({ limit: "maxSharedStrings", max: limits.maxSharedStrings, actual: unique })
      }
      strings = []
      // progress: each string consumes at least 3 bytes, or the loop ends
      for (let k = 0; k < unique; k++) {
        const value = readRichString(segments, limits.maxStringLength)
        if (value === null) {
          return tooLong(limits)
        }
        if (value === undefined) break // fewer strings than declared: keep those read
        strings.push(value)
      }
      continue
    }
    if (type === SUPBOOK && biff8) {
      const supbook = reader.nextWithContinues(limits.maxDecompressedBytes, "maxDecompressedBytes")
      if (supbook === undefined) return malformed("the workbook globals have no EOF")
      if (supbook instanceof MalformedFile || supbook instanceof LimitExceeded) return supbook
      const book = readSupbook(supbook, externalBooks + 1, limits.maxStringLength)
      if (book === null) return tooLong(limits)
      if (book.kind === "external") externalBooks++
      books.push(book)
      continue
    }
    const record = reader.next()
    if (record === undefined) return malformed("the workbook globals have no EOF")
    if (record instanceof MalformedFile) return record
    const f = new Fields(record.data)
    switch (record.type) {
      case EOF:
        break
      case FILEPASS:
        return new Encrypted({ format: "xls" })
      case CODEPAGE: {
        const cp = f.u16(0)
        if (cp !== undefined) text = { ...text, decoder: textDecoder(codepageLabel(cp)) }
        continue
      }
      case DATEMODE:
        date1904 = f.u16(0) === 1
        continue
      case BOUNDSHEET: {
        if (sheets.length >= limits.maxSheets) {
          return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: sheets.length + 1 })
        }
        const offset = f.u32(0)
        const state = f.u8(4)
        const dt = f.u8(5)
        const name = readString(record.data, 6, text, 1)
        if (offset === undefined || state === undefined || dt === undefined || name === undefined) {
          return malformed("truncated BOUNDSHEET", record.offset)
        }
        if (name === null) return tooLong(limits)
        if (dt === 6) {
          modules.push(offset) // a VB module, not a sheet
          continue
        }
        sheets.push({
          name: name.value,
          visibility: (state & 0x03) === 1 ? "hidden" : (state & 0x03) === 2 ? "veryHidden" : "visible",
          offset,
          kind: dt === 0 ? "worksheet" : dt === 1 ? "macrosheet" : dt === 2 ? "chartsheet" : "other"
        })
        continue
      }
      case FORMAT: {
        const id = f.u16(0)
        const code = readString(record.data, 2, text, biff8 ? 2 : 1)
        if (code === null) return tooLong(limits)
        if (id !== undefined && code !== undefined) custom.set(id, code.value)
        continue
      }
      case XF:
        xfFormats.push(f.u16(2) ?? 0)
        continue
      case EXTERNSHEET: {
        if (!biff8) continue
        const count = f.u16(0) ?? 0
        xtis = []
        // progress: k up to count; each entry is 6 bytes inside the record
        for (let k = 0; k < count && 2 + 6 * k + 6 <= record.data.length; k++) {
          xtis.push({ book: f.u16(2 + 6 * k)!, first: f.u16(2 + 6 * k + 2)!, last: f.u16(2 + 6 * k + 4)! })
        }
        continue
      }
      case NAME: {
        if (!biff8) continue
        const grbit = f.u16(0)
        const cch = f.u8(3)
        const cce = f.u16(4)
        const itab = f.u16(8)
        const flags = f.u8(14)
        if (grbit === undefined || cch === undefined || cce === undefined || itab === undefined || flags === undefined) continue
        if (cch > limits.maxStringLength) return tooLong(limits, cch)
        const wide = (flags & 0x01) !== 0
        const nameEnd = 15 + cch * (wide ? 2 : 1)
        if (nameEnd + cce > record.data.length) continue
        if (names.length >= limits.maxDefinedNames) {
          return new LimitExceeded({ limit: "maxDefinedNames", max: limits.maxDefinedNames, actual: names.length + 1 })
        }
        const raw = record.data.subarray(15, nameEnd)
        const chars = wide ? utf16.decode(raw) : String.fromCharCode(...raw)
        const builtin = (grbit & 0x20) !== 0
        const name = builtin ? `_xlnm.${BUILTIN_NAMES[chars.charCodeAt(0)] ?? chars}` : chars
        names.push({ name, hidden: (grbit & 0x01) !== 0, scope: itab > 0 ? itab - 1 : undefined, rgce: record.data.subarray(nameEnd, nameEnd + cce) })
        continue
      }
      default:
        continue
    }
    break // EOF
  }

  const ended = sheetEnds(stream, reader.offset, sheets, modules)
  if (ended instanceof MalformedFile) return ended
  return {
    biff8,
    text,
    date1904,
    sheets: ended,
    formats: xfFormats.map((id) => formatFor(id, custom)),
    strings,
    definedNames: names.map((n) => ({
      name: n.name,
      formula: nameFormula(n.rgce, xtis, books, ended),
      hidden: n.hidden,
      scope: n.scope !== undefined && n.scope < sheets.length ? n.scope : undefined
    }))
  }
}

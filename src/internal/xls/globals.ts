/**
 * The workbook globals substream of a BIFF8 or BIFF5 workbook stream: sheet
 * list (BOUNDSHEET), code page, date system, number formats (FORMAT) and
 * cell styles (XF), the shared-string table (SST with CONTINUE), defined
 * names (NAME with EXTERNSHEET) and encryption (FILEPASS).
 *
 * The walk reads records from the first BOF to its EOF; every record
 * advances the reader. Limits: sheets, shared strings, string lengths, and
 * the SST's total size (maxDecompressedBytes).
 *
 * @internal
 */
import { Encrypted, LimitExceeded, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { EOF, RecordReader, readBof } from "../biff/records.ts"
import { readRichString, SegmentReader } from "../biff/segments.ts"
import { formatFor, type NumberFormat } from "../xlsx/numFmt.ts"
import { codepageLabel, Fields, readString, type Text } from "./fields.ts"

const BOUNDSHEET = 0x0085
const CODEPAGE = 0x0042
const DATEMODE = 0x0022
const FORMAT = 0x041E
const XF = 0x00E0
const SST = 0x00FC
const FILEPASS = 0x002F
const NAME = 0x0018
const EXTERNSHEET = 0x0017

export type SheetKind = "worksheet" | "macrosheet" | "chartsheet" | "other"

export interface SheetEntry {
  readonly name: string
  readonly visibility: "visible" | "hidden" | "veryHidden"
  /** Offset of the sheet's BOF in the workbook stream. */
  readonly offset: number
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

const malformed = (reason: string, offset?: number) =>
  new MalformedFile({ format: "xls", part: "workbook globals", reason, ...(offset === undefined ? {} : { offset }) })

/** Built-in defined names (NAME with fBuiltin): one character code each (MS-XLS 2.5.164 Lbl). */
const BUILTIN_NAMES = [
  "Consolidate_Area", "Auto_Open", "Auto_Close", "Extract", "Database", "Criteria", "Print_Area",
  "Print_Titles", "Recorder", "Data_Form", "Auto_Activate", "Auto_Deactivate", "Sheet_Title", "_FilterDatabase"
]

const quoteSheet = (name: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)

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

/** A cell reference from a BIFF8 row and column word (bit 14: column relative, bit 15: row relative). */
const ref = (row: number, col: number): string =>
  `${(col & 0x4000) !== 0 ? "" : "$"}${column(col & 0x00FF)}${(col & 0x8000) !== 0 ? "" : "$"}${row + 1}`

/**
 * The formula of a defined name, for the forms names usually take: a
 * single 3-D reference or area (`Sheet!$A$1`, `Sheet!$A$1:$B$2`). Anything
 * else gives `""` (formulas aren't decompiled).
 */
const nameFormula = (
  rgce: Uint8Array,
  externSheets: ReadonlyArray<number>,
  sheets: ReadonlyArray<SheetEntry>
): string => {
  const f = new Fields(rgce)
  const ptg = f.u8(0)
  const sheet = (ixti: number | undefined) => {
    const itab = ixti === undefined ? undefined : externSheets[ixti]
    const entry = itab === undefined ? undefined : sheets[itab]
    return entry === undefined ? undefined : quoteSheet(entry.name)
  }
  if ((ptg === 0x3A || ptg === 0x5A || ptg === 0x7A) && rgce.length === 7) {
    const name = sheet(f.u16(1))
    const row = f.u16(3)
    const col = f.u16(5)
    if (name !== undefined && row !== undefined && col !== undefined) return `${name}!${ref(row, col)}`
  }
  if ((ptg === 0x3B || ptg === 0x5B || ptg === 0x7B) && rgce.length === 11) {
    const name = sheet(f.u16(1))
    const [r1, r2, c1, c2] = [f.u16(3), f.u16(5), f.u16(7), f.u16(9)]
    if (name !== undefined && r1 !== undefined && r2 !== undefined && c1 !== undefined && c2 !== undefined) {
      return `${name}!${ref(r1, c1)}:${ref(r2, c2)}`
    }
  }
  return ""
}

/** Reads the globals of a workbook stream (its first substream). */
export const readGlobals = (stream: Uint8Array, limits: Limits): Globals | SpreadsheetError => {
  const reader = new RecordReader(stream, "workbook globals")
  const first = reader.next()
  if (first === undefined) return malformed("empty workbook stream")
  if (first instanceof MalformedFile) return first
  const bof = readBof(first)
  if (bof === undefined || bof.kind !== 0x0005) return malformed("the workbook stream doesn't start with a workbook BOF", 0)
  if (bof.version !== 0x0600 && bof.version !== 0x0500) {
    return new UnsupportedFormat({ format: "xls", reason: `BIFF version 0x${bof.version.toString(16)} in a BIFF5/8 BOF` })
  }
  const biff8 = bof.version === 0x0600
  let text: Text = { biff8, decoder: new TextDecoder("windows-1252"), maxLength: limits.maxStringLength }
  let date1904 = false
  const sheets: Array<SheetEntry> = []
  const custom = new Map<number, string>()
  const xfFormats: Array<number> = []
  let strings: Array<string> = []
  const names: Array<{ name: string; hidden: boolean; scope: number | undefined; rgce: Uint8Array }> = []
  let externSheets: Array<number> = []

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
          return new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual: limits.maxStringLength + 1 })
        }
        if (value === undefined) break // fewer strings than declared: keep those read
        strings.push(value)
      }
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
        if (cp !== undefined) text = { ...text, decoder: new TextDecoder(codepageLabel(cp)) }
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
        if (name === null) return new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual: limits.maxStringLength + 1 })
        if (dt === 6) continue // a VB module, not a sheet
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
        if (id !== undefined && code !== undefined && code !== null) custom.set(id, code.value)
        continue
      }
      case XF:
        xfFormats.push(f.u16(2) ?? 0)
        continue
      case EXTERNSHEET: {
        if (!biff8) continue
        const count = f.u16(0) ?? 0
        externSheets = []
        // progress: k up to count; each entry is 6 bytes inside the record
        for (let k = 0; k < count && 2 + 6 * k + 6 <= record.data.length; k++) externSheets.push(f.u16(2 + 6 * k + 2)!)
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
        const wide = (flags & 0x01) !== 0
        const nameEnd = 15 + cch * (wide ? 2 : 1)
        if (nameEnd + cce > record.data.length) continue
        const raw = record.data.subarray(15, nameEnd)
        const chars = wide ? new TextDecoder("utf-16le").decode(raw) : String.fromCharCode(...raw)
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

  return {
    biff8,
    text,
    date1904,
    sheets,
    formats: xfFormats.map((id) => formatFor(id, custom)),
    strings,
    definedNames: names.map((n) => ({
      name: n.name,
      formula: nameFormula(n.rgce, externSheets, sheets),
      hidden: n.hidden,
      scope: n.scope !== undefined && n.scope < sheets.length ? n.scope : undefined
    }))
  }
}

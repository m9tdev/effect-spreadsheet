/**
 * The workbook globals substream of a BIFF8 or BIFF5 workbook stream: sheet
 * list (BOUNDSHEET), code page, date system, number formats (FORMAT) and
 * cell styles (XF), the shared-string table (SST with CONTINUE), defined
 * names (NAME with EXTERNSHEET and SUPBOOK) and encryption (FILEPASS).
 *
 * The walk reads records from the first BOF to its EOF; every record
 * advances the reader. It goes in batches of at most `BATCH_RECORDS`
 * records, CONTINUEs or strings and about `BATCH_BYTES` bytes, with a yield
 * to the event loop between them, so a file that is nearly all globals
 * (format, style or name floods, a huge SST) doesn't block it. Limits:
 * sheets (and VB modules), defined names, shared strings, string lengths,
 * and the SST's and each SUPBOOK's total size (maxDecompressedBytes).
 *
 * @internal
 */
import { absurd, Effect } from "effect"
import { Encrypted, LimitExceeded, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { DefinedName } from "../../Workbook.ts"
import { readRichString, RecordCursor } from "../biff/cursor.ts"
import { type BiffRecord, BOF_BIFF5_8, EOF, RecordReader, readBof } from "../biff/records.ts"
import { FixedReader } from "../binary/layout.ts"
import { WorkMeter } from "../meter.ts"
import { addFormat, cellStyles, type CellStyles, type NumberFormat } from "../numFmt.ts"
import { areaText, BIFF8_GRID, cellText, sheetPrefix } from "../refText.ts"
import { uniqueSheetNames } from "../sheetNames.ts"
import { type DefinedNameEntry, toDefinedNames } from "../workbook.ts"
import { yieldNow } from "../yield.ts"
import { codepageLabel, readString, type Text, textDecoder } from "./fields.ts"
import { BATCH_CELLS, BATCH_RECORDS } from "./sheet.ts"

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

/*
 * The declared minimum lengths of the fixed fields each record is read
 * with (MS-XLS 2.4): a record is checked once against its minimum, then its
 * fields are read without a check each. Strings after them check their own.
 */
/** CODEPAGE, DATEMODE: one u16. */
const U16_RECORD_MIN = 2
/** BOUNDSHEET: lbPlyPos u32, hsState u8, dt u8, then the name. */
const BOUNDSHEET_MIN = 6
/** FORMAT: ifmt u16, then the format code. */
const FORMAT_MIN = 2
/** XF: ifnt u16, ifmt u16 (the rest isn't read). */
const XF_MIN = 4
/** SUPBOOK: ctab u16, then a marker u16 and nothing else (this workbook, add-ins), or the path and sheet names. */
const SUPBOOK_CTAB = 2
const SUPBOOK_MARKER_LENGTH = 4
/** SST: cstTotal u32, cstUnique u32, then the strings. */
const SST_HEADER = 8
/** EXTERNSHEET: cXTI u16, then 6-byte XTIs (iSupBook, itabFirst, itabLast: u16 each). */
const EXTERNSHEET_COUNT = 2
const XTI_SIZE = 6
/**
 * NAME: grbit u16, chKey u8, cch u8, cce u16, reserved u16, itab u16,
 * 4 reserved bytes, then the name's option byte (its characters and the
 * formula follow).
 */
const NAME_HEADER = 15
/** A name formula of one PtgRef3d / PtgRefErr3d: ptg u8, ixti u16, row u16, column u16. */
const REF3D_LENGTH = 7
/** A name formula of one PtgArea3d / PtgAreaErr3d: ptg u8, ixti u16, rows u16 x2, columns u16 x2. */
const AREA3D_LENGTH = 11

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

export interface Globals {
  readonly biff8: boolean
  readonly text: Text
  readonly date1904: boolean
  readonly sheets: ReadonlyArray<SheetEntry>
  /** Number formats by XF index (a cell's `ixfe`). */
  readonly styles: CellStyles
  readonly strings: ReadonlyArray<string>
  /** The defined names (`scope` only when the name is local to a sheet). */
  readonly definedNames: ReadonlyArray<DefinedName>
}

/** A string longer than maxStringLength: `actual` is its character count. */
const tooLong = (limits: Limits, actual: number) => new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual })

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
  reader: FixedReader,
  rgce: Uint8Array,
  xtis: ReadonlyArray<Xti>,
  books: ReadonlyArray<Book>,
  sheetNames: ReadonlyArray<string>
): string => {
  const prefix = (ixti: number): string | undefined => {
    const xti = xtis[ixti]
    if (xti === undefined) return undefined
    const book = books.length === 0 && xti.book === 0 ? SELF : books[xti.book]
    if (book === undefined || book.kind === "addin") return "#REF!"
    const names = book.kind === "self" ? sheetNames : book.sheets
    const first = names[xti.first]
    const last = names[xti.last]
    if (first === undefined || last === undefined || xti.last < xti.first) return "#REF!"
    return sheetPrefix(book.kind === "self" ? "" : `[${book.link}]`, xti.first === xti.last ? [first] : [first, last])
  }
  // only a formula of exactly one of these tokens is read: its length is the token's
  const f = reader.fixed(rgce, REF3D_LENGTH)
  if (f === undefined) return ""
  const ptg = f.u8(0)
  if ((ptg === 0x3A || ptg === 0x5A || ptg === 0x7A) && rgce.length === REF3D_LENGTH) {
    const sheet = prefix(f.u16(1))
    return sheet === undefined ? "" : sheet + cellText(f.u16(3), f.u16(5))
  }
  if ((ptg === 0x3B || ptg === 0x5B || ptg === 0x7B) && rgce.length === AREA3D_LENGTH) {
    const sheet = prefix(f.u16(1))
    return sheet === undefined ? "" : sheet + areaText(f.u16(3), f.u16(5), f.u16(7), f.u16(9), BIFF8_GRID)
  }
  // PtgRefErr3d / PtgAreaErr3d: a reference whose cells were deleted
  if (
    ((ptg === 0x3C || ptg === 0x5C || ptg === 0x7C) && rgce.length === REF3D_LENGTH) ||
    ((ptg === 0x3D || ptg === 0x5D || ptg === 0x7D) && rgce.length === AREA3D_LENGTH)
  ) {
    const sheet = prefix(f.u16(1))
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

/**
 * Bytes of record data per batch (see `GlobalsWalk.step`): records of up to
 * 8 224 bytes each, and an SST or SUPBOOK of up to `maxDecompressedBytes`,
 * would otherwise make a batch of `BATCH_RECORDS` records or strings cost
 * megabytes of decoding.
 */
export const BATCH_BYTES = 16 * BATCH_CELLS

/** A record the walk reads with its CONTINUEs. */
type ContinuedKind = "Sst" | "Supbook" | "ExternSheet" | "Name"

/**
 * What the walk works on: plain records; or the CONTINUEs of an SST,
 * SUPBOOK, EXTERNSHEET or NAME (`record`, the last one read), checked and
 * passed over; then the strings of an SST or SUPBOOK, read through a
 * cursor over its data and its CONTINUEs'.
 */
type Pending =
  | { readonly _tag: "Records" }
  | { readonly _tag: "Continues"; readonly of: ContinuedKind; readonly record: BiffRecord }
  | { readonly _tag: "SstStrings"; readonly segments: RecordCursor; readonly unique: number }
  | {
    readonly _tag: "SupbookSheets"
    readonly segments: RecordCursor
    readonly ctab: number
    readonly sheets: Array<string>
  }

const RECORDS: Pending = { _tag: "Records" }

type ContinuesPending = Extract<Pending, { readonly _tag: "Continues" }>
type StringsPending = Extract<Pending, { readonly _tag: "SstStrings" | "SupbookSheets" }>

/**
 * The walk over the globals substream, from the first BOF to its EOF, in
 * bounded steps (`step`). Every unit of work (a record, a CONTINUE, a
 * string) advances the record reader or a string reader.
 */
class GlobalsWalk {
  readonly stream: Uint8Array
  readonly reader: RecordReader
  readonly limits: Limits
  biff8 = true
  text: Text
  date1904 = false
  readonly sheets: Array<Omit<SheetEntry, "end">> = []
  /** Offsets of VB module BOUNDSHEETs: not sheets, but they may start a substream. */
  readonly modules: Array<number> = []
  /** Custom number formats by id, classified as their FORMAT is read. */
  readonly custom = new Map<number, NumberFormat>()
  readonly xfFormats: Array<number> = []
  strings: Array<string> = []
  readonly names: Array<{ name: string; hidden: boolean; scope: number | undefined; rgce: Uint8Array }> = []
  readonly books: Array<Book> = []
  externalBooks = 0
  xtis: Array<Xti> = []
  pending: Pending = RECORDS
  /** Pointed at each record whose fixed fields are read. */
  private readonly fields = new FixedReader()
  /** Work done in the current batch: records, CONTINUEs and strings read, and their bytes. */
  private units = 0
  private bytes = 0

  constructor(stream: Uint8Array, limits: Limits) {
    this.stream = stream
    this.reader = new RecordReader(stream, "workbook globals")
    this.limits = limits
    this.text = { biff8: true, decoder: textDecoder("windows-1252"), maxLength: limits.maxStringLength }
  }

  /** Reads the workbook BOF. */
  start(): SpreadsheetError | undefined {
    const first = this.reader.next()
    if (first === undefined) return malformed("empty workbook stream")
    if (first instanceof MalformedFile) return first
    const bof = readBof(first)
    if (bof === undefined || bof.kind !== 0x0005) return malformed("the workbook stream doesn't start with a workbook BOF", 0)
    // BIFF5 writers put 0x0500 or another 0x05xx (e.g. 0x0580) here
    if (bof.version !== 0x0600 && (bof.version & 0xFF00) !== 0x0500) {
      return new UnsupportedFormat({ format: "xls", reason: `BIFF version 0x${bof.version.toString(16)} in a BIFF5/8 BOF` })
    }
    this.biff8 = bof.version === 0x0600
    this.text = { ...this.text, biff8: this.biff8 }
    return undefined
  }

  /**
   * One batch: records (CONTINUEs included) and strings until `BATCH_RECORDS`
   * of them or `BATCH_BYTES` of their data. "done" at the EOF.
   */
  step(): "more" | "done" | SpreadsheetError {
    this.units = 0
    this.bytes = 0
    // progress: every iteration reads at least one record, CONTINUE or string
    // (each advances a reader), and adds it to units, or returns
    while (this.units < BATCH_RECORDS && this.bytes < BATCH_BYTES) {
      const pending = this.pending
      let outcome: "done" | SpreadsheetError | undefined
      switch (pending._tag) {
        case "Records":
          outcome = this.record()
          break
        case "Continues":
          outcome = this.continues(pending)
          break
        case "SstStrings":
        case "SupbookSheets":
          outcome = this.readStrings(pending)
          break
        default:
          return absurd(pending)
      }
      if (outcome !== undefined) return outcome
    }
    return "more"
  }

  /** Reads one record (an SST, SUPBOOK, EXTERNSHEET or NAME becomes `pending`); "done" at the EOF. */
  private record(): "done" | SpreadsheetError | undefined {
    const { reader, limits, biff8 } = this
    const record = reader.next()
    if (record === undefined) return malformed("the workbook globals have no EOF")
    if (record instanceof MalformedFile) return record
    this.units++
    this.bytes += 4 + record.data.length
    const { data } = record
    switch (record.type) {
      case EOF:
        return "done"
      case FILEPASS:
        return new Encrypted({ format: "xls" })
      // BIFF5 has no SST; its EXTERNSHEETs and NAMEs aren't read, so neither are its SUPBOOKs
      case SST:
        return biff8 ? this.startContinues(record, "Sst") : undefined
      case SUPBOOK:
        return biff8 ? this.startContinues(record, "Supbook") : undefined
      case EXTERNSHEET:
        return biff8 ? this.startContinues(record, "ExternSheet") : undefined
      case NAME:
        return biff8 ? this.startContinues(record, "Name") : undefined
      case CODEPAGE: {
        // a cut CODEPAGE is ignored
        const f = this.fields.fixed(data, U16_RECORD_MIN)
        if (f !== undefined) this.text = { ...this.text, decoder: textDecoder(codepageLabel(f.u16(0))) }
        return undefined
      }
      case DATEMODE: {
        // a cut DATEMODE means the 1900 system
        const f = this.fields.fixed(data, U16_RECORD_MIN)
        this.date1904 = f !== undefined && f.u16(0) === 1
        return undefined
      }
      case BOUNDSHEET:
        return this.boundSheet(record)
      case FORMAT: {
        // a cut FORMAT is ignored; a code over the limit is not
        const f = this.fields.fixed(data, FORMAT_MIN)
        if (f === undefined) return undefined
        const id = f.u16(0)
        const code = readString(data, FORMAT_MIN, this.text, biff8 ? 2 : 1)
        if (code._tag === "TooLong") return tooLong(limits, code.count)
        if (code._tag === "Ok") {
          // classified here (a later FORMAT with the same id wins); the
          // scan counts in the batch, by the code's length
          addFormat(this.custom, id, code.value)
          this.bytes += code.value.length
        }
        return undefined
      }
      case XF: {
        // a cut XF still takes its index, with the General format
        const f = this.fields.fixed(data, XF_MIN)
        this.xfFormats.push(f === undefined ? 0 : f.u16(2))
        return undefined
      }
      default:
        return undefined
    }
  }

  /** A BOUNDSHEET: a sheet, or a VB module's offset. A cut one is MalformedFile. */
  private boundSheet(record: BiffRecord): SpreadsheetError | undefined {
    const { limits } = this
    const f = this.fields.fixed(record.data, BOUNDSHEET_MIN)
    if (f === undefined) return malformed("truncated BOUNDSHEET", record.offset)
    const offset = f.u32(0)
    const state = f.u8(4)
    const dt = f.u8(5)
    const name = readString(record.data, BOUNDSHEET_MIN, this.text, 1)
    if (name._tag === "TooLong") return tooLong(limits, name.count)
    if (name._tag === "Truncated") return malformed("truncated BOUNDSHEET", record.offset)
    if (dt === 6) {
      // a VB module, not a sheet; real workbooks have a few, so they share maxSheets' bound
      if (this.modules.length >= limits.maxSheets) {
        return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: this.modules.length + 1 })
      }
      this.modules.push(offset)
      return undefined
    }
    if (this.sheets.length >= limits.maxSheets) {
      return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: this.sheets.length + 1 })
    }
    this.sheets.push({
      name: name.value,
      visibility: (state & 0x03) === 1 ? "hidden" : (state & 0x03) === 2 ? "veryHidden" : "visible",
      offset,
      kind: dt === 0 ? "worksheet" : dt === 1 ? "macrosheet" : dt === 2 ? "chartsheet" : "other"
    })
    return undefined
  }

  /**
   * An EXTERNSHEET (with its CONTINUEs): the XTIs, as many as declared and
   * present. A few hundred kilobytes at most (65 535 entries of 6 bytes),
   * all counted in the batch's bytes while its CONTINUEs were read.
   */
  private externSheet(segments: RecordCursor): undefined {
    const count = segments.has(EXTERNSHEET_COUNT) ? segments.u16() : 0
    this.xtis = []
    // progress: k up to count; each entry reads 6 of the bytes left
    for (let k = 0; k < count && segments.has(XTI_SIZE); k++) {
      this.xtis.push({ book: segments.u16(), first: segments.u16(), last: segments.u16() })
    }
    return undefined
  }

  /**
   * A NAME (with its CONTINUEs): its name and formula. One whose header,
   * name or formula is cut short is skipped. Name characters continued in a
   * CONTINUE start with their own option byte there, as other strings do.
   */
  private name(segments: RecordCursor): LimitExceeded | undefined {
    const { limits } = this
    if (!segments.has(NAME_HEADER)) return undefined
    const grbit = segments.u16()
    segments.u8() // keyboard shortcut
    const cch = segments.u8()
    const cce = segments.u16()
    segments.skip(2)
    const itab = segments.u16()
    segments.skip(4)
    const flags = segments.u8()
    if (cch > limits.maxStringLength) return tooLong(limits, cch)
    const chars = segments.chars(cch, (flags & 0x01) !== 0)
    if (chars._tag !== "Ok" || !segments.has(cce)) return undefined
    if (this.names.length >= limits.maxDefinedNames) {
      return new LimitExceeded({ limit: "maxDefinedNames", max: limits.maxDefinedNames, actual: this.names.length + 1 })
    }
    const rgce = new Uint8Array(cce)
    // progress: k up to cce, with at least cce bytes left
    for (let k = 0; k < cce; k++) rgce[k] = segments.u8()
    const builtin = (grbit & 0x20) !== 0
    const text = chars.value
    const name = builtin ? `_xlnm.${BUILTIN_NAMES[text.charCodeAt(0)] ?? text}` : text
    this.names.push({ name, hidden: (grbit & 0x01) !== 0, scope: itab > 0 ? itab - 1 : undefined, rgce })
    return undefined
  }

  /** Starts reading `record`'s CONTINUEs (the next steps go on with them). */
  private startContinues(record: BiffRecord, of: ContinuedKind): undefined {
    this.pending = { _tag: "Continues", of, record }
    return undefined
  }

  /**
   * Goes on with a record that may be continued: checks its CONTINUEs and
   * moves past them within the batch's budget (all of them together within
   * maxDecompressedBytes), then reads an EXTERNSHEET or NAME whole, or
   * starts on an SST's or SUPBOOK's strings.
   */
  private continues(pending: ContinuesPending): SpreadsheetError | undefined {
    const { reader, limits } = this
    const { record } = pending
    const before = reader.records
    const at = reader.offset
    const step = reader.continues(record, BATCH_RECORDS - this.units, limits.maxDecompressedBytes, "maxDecompressedBytes")
    this.units += reader.records - before
    this.bytes += reader.offset - at
    if (step === "more") return undefined
    if (step !== "done") return step
    // the CONTINUEs end where the reader is: the cursor reads up to there
    const segments = new RecordCursor(this.stream, record, reader.offset)
    switch (pending.of) {
      case "ExternSheet":
        this.pending = RECORDS
        return this.externSheet(segments)
      case "Name":
        this.pending = RECORDS
        return this.name(segments)
      case "Sst": {
        if (!segments.has(SST_HEADER)) return malformed("truncated SST", record.offset)
        segments.u32() // total count
        const unique = segments.u32()
        if (unique > limits.maxSharedStrings) {
          return new LimitExceeded({ limit: "maxSharedStrings", max: limits.maxSharedStrings, actual: unique })
        }
        this.strings = []
        this.pending = { _tag: "SstStrings", segments, unique }
        return undefined
      }
      case "Supbook": {
        // this workbook and add-ins are exactly ctab + marker (4 bytes); an
        // external book's path can have a character count equal to a marker
        const head = record.data.length === SUPBOOK_MARKER_LENGTH ? this.fields.fixed(record.data, SUPBOOK_MARKER_LENGTH) : undefined
        const marker = head?.u16(2)
        if (marker === 0x0401 || marker === 0x3A01) {
          this.books.push(marker === 0x0401 ? SELF : { kind: "addin" })
          this.pending = RECORDS
          return undefined
        }
        const ctab = segments.has(SUPBOOK_CTAB) ? segments.u16() : 0
        const path = readRichString(segments, limits.maxStringLength) // the path's character count is the marker's place
        if (path._tag === "TooLong") return tooLong(limits, path.count)
        this.units++
        // without a path, no sheet names follow
        this.pending = { _tag: "SupbookSheets", segments, ctab: path._tag === "Truncated" ? 0 : ctab, sheets: [] }
        return undefined
      }
      default:
        return absurd(pending.of)
    }
  }

  /** Reads an SST's strings or a SUPBOOK's sheet names within the batch's budget. */
  private readStrings(pending: StringsPending): SpreadsheetError | undefined {
    const { limits } = this
    const { segments } = pending
    let list: Array<string>
    let count: number
    switch (pending._tag) {
      case "SstStrings":
        list = this.strings
        count = pending.unique
        break
      case "SupbookSheets":
        list = pending.sheets
        count = pending.ctab
        break
      default:
        return absurd(pending)
    }
    const started = segments.consumed()
    let ended = false
    // progress: each string consumes at least 3 bytes, or the list ends; at most the batch's units
    while (this.units < BATCH_RECORDS && segments.consumed() - started < BATCH_BYTES - this.bytes) {
      if (list.length >= count) {
        ended = true
        break
      }
      const read = readRichString(segments, limits.maxStringLength)
      if (read._tag === "TooLong") return tooLong(limits, read.count)
      this.units++
      if (read._tag === "Truncated") {
        ended = true // fewer strings than declared: keep those read
        break
      }
      list.push(read.value)
    }
    this.bytes += segments.consumed() - started
    if (!ended && list.length < count) return undefined
    this.pending = RECORDS
    switch (pending._tag) {
      case "SstStrings":
        return undefined
      case "SupbookSheets":
        this.externalBooks++
        this.books.push({ kind: "external", link: this.externalBooks, sheets: pending.sheets })
        return undefined
      default:
        return absurd(pending)
    }
  }
}

/**
 * Reads the globals of a workbook stream (its first substream), in batches
 * of bounded work with a yield to the event loop between them, and so are
 * its defined names. Sheet names are made unique
 * (`uniqueSheetNames`). Each batch counts one unit of work on the meter.
 */
export const readGlobals = (stream: Uint8Array, limits: Limits): Effect.Effect<Globals, SpreadsheetError> =>
  Effect.gen(function*() {
    const meter = yield* WorkMeter
    const walk = new GlobalsWalk(stream, limits)
    const started = walk.start()
    if (started !== undefined) return yield* Effect.fail(started)
    // progress: every step reads at least one record, CONTINUE or string, or ends
    for (let k = 0;; k++) {
      if (k > 0) yield* yieldNow
      meter.add("batches", 1)
      const outcome = walk.step()
      if (outcome === "done") break
      if (outcome !== "more") return yield* Effect.fail(outcome)
    }
    const { sheets, xtis, books, custom } = walk
    const unique = uniqueSheetNames(sheets.map((s) => s.name))
    const ended = sheetEnds(stream, walk.reader.offset, sheets.map((s, k) => ({ ...s, name: unique[k]! })), walk.modules)
    if (ended instanceof MalformedFile) return yield* Effect.fail(ended)

    // a lookup per cell: the codes were classified as their FORMATs were read
    const styles = cellStyles(walk.xfFormats, walk.xfFormats.length, custom)
    const definedNames: Array<DefinedNameEntry> = []
    const sheetNames = ended.map((s) => s.name) // once, not per name
    const formulas = new FixedReader()
    // progress: k up to names.length (at most maxDefinedNames), yielding every BATCH_RECORDS
    for (let k = 0; k < walk.names.length; k++) {
      if (k % BATCH_RECORDS === 0) {
        if (k > 0) yield* yieldNow
        meter.add("batches", 1)
      }
      const n = walk.names[k]!
      const formula = nameFormula(formulas, n.rgce, xtis, books, sheetNames)
      definedNames.push({ name: n.name, formula, hidden: n.hidden, scope: n.scope })
    }
    return {
      biff8: walk.biff8,
      text: walk.text,
      date1904: walk.date1904,
      sheets: ended,
      styles,
      strings: walk.strings,
      definedNames: toDefinedNames(definedNames, sheets.length)
    }
  })

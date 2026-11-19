/**
 * The workbook globals substream of a BIFF8 or BIFF5 workbook stream: sheet
 * list (BOUNDSHEET), code page, date system, number formats (FORMAT) and
 * cell styles (XF), the shared-string table (SST with CONTINUE), defined
 * names (NAME with EXTERNSHEET and SUPBOOK) and encryption (FILEPASS).
 *
 * The walk reads records from the first BOF to its EOF; every record
 * advances the reader. It is a generator driven by `drive`, which yields to
 * the event loop after at most `BATCH_RECORDS` records, CONTINUEs or
 * strings or about `BATCH_BYTES` bytes of them, so a file that is nearly
 * all globals (format, style or name floods, a huge SST) doesn't block it.
 * Limits: sheets (and VB modules), defined names, shared strings, string
 * lengths, and the SST's and each SUPBOOK's total size
 * (maxDecompressedBytes).
 *
 * @internal
 */
import { absurd, type Effect, Result } from "effect"
import { Encrypted, LimitExceeded, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { DefinedName } from "../../Workbook.ts"
import { readRichString, RecordCursor } from "../biff/cursor.ts"
import { type BiffRecord, BOF_BIFF5_8, CONTINUE, EOF, RecordReader, readBof } from "../biff/records.ts"
import { FixedReader } from "../binary/layout.ts"
import { addFormat, cellStyles, type CellStyles, type NumberFormat } from "../numFmt.ts"
import { areaText, BIFF8_GRID, cellText, sheetPrefix } from "../refText.ts"
import { uniqueSheetNames } from "../sheetNames.ts"
import { drive, type Turn, type Walk } from "../walk.ts"
import { type DefinedNameEntry, toDefinedNames } from "../workbook.ts"
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
 * The walk's budget per turn (see `Turn`), in bytes: a record, CONTINUE,
 * string or defined name costs its bytes of data plus `UNIT`, so a turn
 * reads at most `BATCH_RECORDS` of them, or about `BATCH_BYTES` bytes of
 * their data. Records of up to 8 224 bytes each, and an SST or SUPBOOK of
 * up to `maxDecompressedBytes`, would otherwise make a turn of
 * `BATCH_RECORDS` of them cost megabytes of decoding.
 */
export const BATCH_BYTES = 16 * BATCH_CELLS

/** What a record, CONTINUE, string or defined name costs besides its bytes: `BATCH_RECORDS` of them fill a turn. */
const UNIT = BATCH_BYTES / BATCH_RECORDS

/** A record the walk reads with its CONTINUEs. */
type ContinuedKind = "Sst" | "Supbook" | "ExternSheet" | "Name"

/** Part of the walk (see `Walk`): yields when the turn's budget is spent. */
type Steps<A> = Generator<void, A, void>

/**
 * What the globals walk gathers, and the steps it takes over the globals
 * substream, from the first BOF to its EOF (`records`). Every step (a
 * record, a run of CONTINUEs, a string) advances the record reader or a
 * string reader, and counts its cost on the turn.
 */
class GlobalsWalk {
  readonly stream: Uint8Array
  readonly reader: RecordReader
  readonly limits: Limits
  readonly turn: Turn
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
  /** Pointed at each record whose fixed fields are read. */
  private readonly fields = new FixedReader()

  constructor(stream: Uint8Array, limits: Limits, turn: Turn) {
    this.stream = stream
    this.reader = new RecordReader(stream, "workbook globals")
    this.limits = limits
    this.turn = turn
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
   * The records after the BOF, up to the EOF: one step each, and an SST,
   * SUPBOOK, EXTERNSHEET or NAME goes on with its CONTINUEs and then its
   * strings or entries.
   */
  *records(): Steps<SpreadsheetError | undefined> {
    const { reader, biff8, turn } = this
    // progress: every iteration reads a record (the reader advances) and
    // goes on past its CONTINUEs, or returns
    for (;;) {
      const record = reader.next()
      if (record === undefined) return malformed("the workbook globals have no EOF")
      if (record instanceof MalformedFile) return record
      let cost = UNIT + 4 + record.data.length
      let continued: ContinuedKind | undefined
      switch (record.type) {
        case EOF:
          return undefined
        case FILEPASS:
          return new Encrypted({ format: "xls" })
        // BIFF5 has no SST; its EXTERNSHEETs and NAMEs aren't read, so neither are its SUPBOOKs
        case SST:
          continued = "Sst"
          break
        case SUPBOOK:
          continued = "Supbook"
          break
        case EXTERNSHEET:
          continued = "ExternSheet"
          break
        case NAME:
          continued = "Name"
          break
        default: {
          const read = this.record(record)
          if (typeof read !== "number") return read
          cost += read
        }
      }
      if (turn.spend(cost)) yield
      if (continued === undefined || !biff8) continue
      // most of these records have no CONTINUE: their cursor is made at once
      const segments = reader.peekType() === CONTINUE ? yield* this.continues(record) : new RecordCursor(this.stream, record, reader.offset)
      if (!(segments instanceof RecordCursor)) return segments
      // then an EXTERNSHEET or NAME is read whole, an SST's or SUPBOOK's strings in steps
      let failed: SpreadsheetError | undefined
      switch (continued) {
        case "ExternSheet":
          this.externSheet(segments)
          break
        case "Name":
          failed = this.name(segments)
          break
        case "Sst":
          failed = yield* this.sst(record, segments)
          break
        case "Supbook":
          failed = yield* this.supbook(record, segments)
          break
        default:
          return absurd(continued)
      }
      if (failed !== undefined) return failed
    }
  }

  /**
   * Reads a record that isn't continued: its cost besides the record's own
   * (a FORMAT's code is scanned), or an error.
   */
  private record(record: BiffRecord): number | SpreadsheetError {
    const { limits, biff8 } = this
    const { data } = record
    switch (record.type) {
      case CODEPAGE: {
        // a cut CODEPAGE is ignored
        const f = this.fields.fixed(data, U16_RECORD_MIN)
        if (f !== undefined) this.text = { ...this.text, decoder: textDecoder(codepageLabel(f.u16(0))) }
        return 0
      }
      case DATEMODE: {
        // a cut DATEMODE means the 1900 system
        const f = this.fields.fixed(data, U16_RECORD_MIN)
        this.date1904 = f !== undefined && f.u16(0) === 1
        return 0
      }
      case BOUNDSHEET:
        return this.boundSheet(record) ?? 0
      case FORMAT: {
        // a cut FORMAT is ignored; a code over the limit is not
        const f = this.fields.fixed(data, FORMAT_MIN)
        if (f === undefined) return 0
        const id = f.u16(0)
        const code = readString(data, FORMAT_MIN, this.text, biff8 ? 2 : 1)
        if (code._tag === "TooLong") return tooLong(limits, code.count)
        if (code._tag !== "Ok") return 0
        // classified here (a later FORMAT with the same id wins); the scan
        // costs the code's length
        addFormat(this.custom, id, code.value)
        return code.value.length
      }
      case XF: {
        // a cut XF still takes its index, with the General format
        const f = this.fields.fixed(data, XF_MIN)
        this.xfFormats.push(f === undefined ? 0 : f.u16(2))
        return 0
      }
      default:
        return 0
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
   * Moves past the CONTINUEs after `record` and checks them (all of them
   * together within maxDecompressedBytes), in runs of as many as the turn
   * has room for (one at least), each one step; then a cursor over the
   * record's data and theirs.
   */
  private *continues(record: BiffRecord): Steps<RecordCursor | SpreadsheetError> {
    const { reader, limits, turn } = this
    // progress: every run moves past at least one CONTINUE (one is next
    // when the walk comes here, and after a run that says "more"), or is
    // the last
    for (;;) {
      const before = reader.records
      const at = reader.offset
      const room = Math.max(1, Math.floor(turn.left / UNIT))
      const run = reader.continues(record, room, limits.maxDecompressedBytes, "maxDecompressedBytes")
      if (run !== "more" && run !== "done") return run
      const read = reader.records - before
      if (turn.spend(read * UNIT + reader.offset - at)) yield
      // the CONTINUEs end where the reader is: the cursor reads up to there
      if (run === "done") return new RecordCursor(this.stream, record, reader.offset)
    }
  }

  /**
   * An EXTERNSHEET (with its CONTINUEs): the XTIs, as many as declared and
   * present. A few hundred kilobytes at most (65 535 entries of 6 bytes),
   * all counted in the cost of its CONTINUEs.
   */
  private externSheet(segments: RecordCursor): void {
    const count = segments.has(EXTERNSHEET_COUNT) ? segments.u16() : 0
    this.xtis = []
    // progress: k up to count; each entry reads 6 of the bytes left
    for (let k = 0; k < count && segments.has(XTI_SIZE); k++) {
      this.xtis.push({ book: segments.u16(), first: segments.u16(), last: segments.u16() })
    }
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

  /** An SST (with its CONTINUEs): its strings, as many as declared and present. */
  private *sst(record: BiffRecord, segments: RecordCursor): Steps<SpreadsheetError | undefined> {
    const { limits } = this
    if (!segments.has(SST_HEADER)) return malformed("truncated SST", record.offset)
    segments.u32() // total count
    const unique = segments.u32()
    if (unique > limits.maxSharedStrings) {
      return new LimitExceeded({ limit: "maxSharedStrings", max: limits.maxSharedStrings, actual: unique })
    }
    this.strings = []
    return yield* this.readStrings(segments, unique, this.strings)
  }

  /**
   * A SUPBOOK (with its CONTINUEs): this workbook, an add-in's functions,
   * or an external workbook with its path and sheet names.
   */
  private *supbook(record: BiffRecord, segments: RecordCursor): Steps<SpreadsheetError | undefined> {
    const { limits } = this
    // this workbook and add-ins are exactly ctab + marker (4 bytes); an
    // external book's path can have a character count equal to a marker
    const head = record.data.length === SUPBOOK_MARKER_LENGTH ? this.fields.fixed(record.data, SUPBOOK_MARKER_LENGTH) : undefined
    const marker = head?.u16(2)
    if (marker === 0x0401 || marker === 0x3A01) {
      this.books.push(marker === 0x0401 ? SELF : { kind: "addin" })
      return undefined
    }
    const ctab = segments.has(SUPBOOK_CTAB) ? segments.u16() : 0
    const path = readRichString(segments, limits.maxStringLength) // the path's character count is the marker's place
    if (path._tag === "TooLong") return tooLong(limits, path.count)
    const sheets: Array<string> = []
    // without a path, no sheet names follow
    const failed = yield* this.readStrings(segments, path._tag === "Truncated" ? 0 : ctab, sheets)
    if (failed !== undefined) return failed
    this.externalBooks++
    this.books.push({ kind: "external", link: this.externalBooks, sheets })
    return undefined
  }

  /**
   * Reads up to `count` strings into `list`, one step each: fewer when the
   * data ends first (those read are kept).
   */
  private *readStrings(segments: RecordCursor, count: number, list: Array<string>): Steps<LimitExceeded | undefined> {
    const { limits, turn } = this
    // progress: list grows by one per iteration, up to count, or the loop ends
    while (list.length < count) {
      const at = segments.consumed()
      const read = readRichString(segments, limits.maxStringLength)
      if (read._tag === "TooLong") return tooLong(limits, read.count)
      if (read._tag === "Truncated") return undefined
      list.push(read.value)
      if (turn.spend(UNIT + segments.consumed() - at)) yield
    }
    return undefined
  }
}

/**
 * The globals walk (see `Walk`): the records up to the EOF, then the sheets'
 * substream ends (sheet names made unique, `uniqueSheetNames`), the cell
 * styles and the defined names, one step each.
 */
function* globalsWalk(stream: Uint8Array, limits: Limits, turn: Turn): Walk<Globals, SpreadsheetError> {
  const walk = new GlobalsWalk(stream, limits, turn)
  const started = walk.start()
  if (started !== undefined) return Result.fail(started)
  const failed = yield* walk.records()
  if (failed !== undefined) return Result.fail(failed)
  const { sheets, xtis, books, custom } = walk
  const unique = uniqueSheetNames(sheets.map((s) => s.name))
  const ended = sheetEnds(stream, walk.reader.offset, sheets.map((s, k) => ({ ...s, name: unique[k]! })), walk.modules)
  if (ended instanceof MalformedFile) return Result.fail(ended)

  // a lookup per cell: the codes were classified as their FORMATs were read
  const styles = cellStyles(walk.xfFormats, walk.xfFormats.length, custom)
  const definedNames: Array<DefinedNameEntry> = []
  const sheetNames = ended.map((s) => s.name) // once, not per name
  const formulas = new FixedReader()
  // progress: k up to names.length (at most maxDefinedNames)
  for (let k = 0; k < walk.names.length; k++) {
    const n = walk.names[k]!
    const formula = nameFormula(formulas, n.rgce, xtis, books, sheetNames)
    definedNames.push({ name: n.name, formula, hidden: n.hidden, scope: n.scope })
    if (turn.spend(UNIT)) yield
  }
  return Result.succeed({
    biff8: walk.biff8,
    text: walk.text,
    date1904: walk.date1904,
    sheets: ended,
    styles,
    strings: walk.strings,
    definedNames: toDefinedNames(definedNames, sheets.length)
  })
}

/**
 * Reads the globals of a workbook stream (its first substream), driven in
 * turns of at most `BATCH_BYTES` of cost with a yield to the event loop
 * between them (`drive`). Each turn counts one batch on the work meter.
 */
export const readGlobals = (stream: Uint8Array, limits: Limits): Effect.Effect<Globals, SpreadsheetError> =>
  drive((turn) => globalsWalk(stream, limits, turn), BATCH_BYTES, "batches")

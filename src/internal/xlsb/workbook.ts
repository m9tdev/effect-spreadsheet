/**
 * The workbook part of an XLSB package (`xl/workbook.bin`): sheets in
 * order, the date system, the external references (books and XTIs) and the
 * defined names, plus the external link parts' sheet names that names
 * point into.
 *
 * A defined name's formula is parsed tokens; like the XLS reader, this one
 * writes out the forms names usually take, a single 3-D reference or area
 * (`Sheet!$A$1`, `'My Sheet'!$1:$2`, `[1]Data!$A:$A`), as Excel writes them
 * in XLSX, and gives `""` for any other formula (not decompiled). Only
 * those forms are kept while reading (their few numbers, not the tokens).
 *
 * @internal
 */
import { Effect } from "effect"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { DefinedName } from "../../Workbook.ts"
import { readNullableWideString, readWideString } from "../biff12/fields.ts"
import { type Biff12Record, type Handled, readRecords } from "../biff12/records.ts"
import { type Fixed, FixedReader, type StringFailure } from "../binary/layout.ts"
import { malformedIn, type SheetEntry, sheetVisibility } from "../package.ts"
import { areaText, BIFF12_GRID, cellText, sheetPrefix } from "../refText.ts"
import { uniqueSheetNames } from "../sheetNames.ts"
import { toDefinedNames } from "../workbook.ts"
import type { Archive } from "../zip/archive.ts"
import type { ZipEntry } from "../zip/directory.ts"

export const malformed = malformedIn("xlsb")

/**
 * The error of a string read that failed: a `LimitExceeded` with the
 * string's length if it is over maxStringLength, else (it runs past its
 * record) a MalformedFile for `reason`.
 */
export const stringError = (read: StringFailure, limits: Limits, reason: string, part: string, offset: number): SpreadsheetError =>
  read._tag === "TooLong"
    ? new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual: read.count })
    : malformed(reason, part, offset)

const BRT_NAME = 0x0027
const BRT_BEGIN_BOOK = 0x0083
const BRT_WB_PROP = 0x0099
const BRT_BUNDLE_SH = 0x009C
const BRT_SUP_BOOK_SRC = 0x0163
const BRT_SUP_SELF = 0x0165
const BRT_SUP_SAME = 0x0166
const BRT_SUP_TABS = 0x0167
const BRT_EXTERN_SHEET = 0x016A
const BRT_SUP_ADDIN = 0x029B

/**
 * A supporting book (BrtSup* records, in order): this workbook, an add-in,
 * or an external workbook, whose relationship leads to its external link
 * part. `link` numbers the external ones from 1, as XLSX does (`[1]`).
 */
export type Book =
  | { readonly kind: "self" }
  | { readonly kind: "addin" }
  | { readonly kind: "external"; readonly link: number; readonly relationshipId: string | undefined }

const SELF_BOOK: Book = { kind: "self" }
const ADDIN_BOOK: Book = { kind: "addin" }

/**
 * Supporting books and XTIs kept: an XTI's book and a token's XTI are
 * 16-bit indexes in practice (a name's ixti is a u16), so entries past
 * 65 536 can't be told apart from missing ones: they aren't kept (a
 * reference to one is `#REF!`, or `""` for a missing XTI), and a flood of
 * tiny BrtSup* or XTI records costs no memory past them.
 */
export const MAX_BOOKS = 65_536
export const MAX_XTIS = 65_536

/**
 * Where a name's XTI points, resolved once after the workbook part:
 * sheets `first`..`last` of this workbook, or of the external book at
 * index `book` (one with a link relationship); `noBook` for an add-in, a
 * book that isn't there or an external one without a link relationship
 * (`#REF!`); `missing` when there is no such XTI (`""`).
 */
export type Xti =
  | { readonly kind: "missing" }
  | { readonly kind: "noBook" }
  | { readonly kind: "self"; readonly first: number; readonly last: number }
  | {
    readonly kind: "external"
    readonly book: number
    readonly link: number
    readonly relationshipId: string
    readonly first: number
    readonly last: number
  }

const MISSING_XTI: Xti = { kind: "missing" }
const NO_BOOK: Xti = { kind: "noBook" }

/** A single 3-D reference or area of a defined name: its XTI, where that points, and its cells. */
export interface NameRef {
  readonly kind: "ref" | "area" | "refErr" | "areaErr"
  readonly ixti: number
  readonly xti: Xti
  readonly r1: number
  readonly r2: number
  readonly c1: number
  readonly c2: number
}

/** A reference as it is read: its XTI is resolved once the whole part is read. */
interface RefBuilder extends Omit<NameRef, "xti"> {
  xti: Xti
}

export interface NameEntry {
  readonly name: string
  readonly scope: number | undefined
  readonly hidden: boolean
  /** The reference, for the forms written out; `undefined` for any other formula. */
  readonly ref: NameRef | undefined
}

export interface WorkbookBin {
  readonly sheets: ReadonlyArray<SheetEntry>
  readonly date1904: boolean
  readonly books: ReadonlyArray<Book>
  /** XTIs as triples: book, first sheet, last sheet (sheets -1 and -2: no sheet). */
  readonly xtis: Int32Array
  /** The defined names, each reference's XTI resolved. */
  readonly names: ReadonlyArray<NameEntry>
}

/** Ptg types of a 3-D reference, area and their deleted-cells forms, in each class (reference, value, array). */
const REF_3D = new Set([0x3A, 0x5A, 0x7A])
const AREA_3D = new Set([0x3B, 0x5B, 0x7B])
const REF_ERR_3D = new Set([0x3C, 0x5C, 0x7C])
const AREA_ERR_3D = new Set([0x3D, 0x5D, 0x7D])

/**
 * The reference of a name formula's tokens (`rgce`, `cce` bytes at `at`,
 * all inside the record: checked by the caller), when it is a single 3-D
 * one. Its fields are read only for those two lengths, so only inside the
 * formula.
 */
const nameRef = (f: Fixed, at: number, cce: number): RefBuilder | undefined => {
  // PtgRef3d: ptg, ixti u16, row u32, column u16 (flags in its high bits)
  if (cce === 9) {
    const ptg = f.u8(at)
    if (!REF_3D.has(ptg) && !REF_ERR_3D.has(ptg)) return undefined
    const r = f.u32(at + 3)
    const c = f.u16(at + 7)
    return { kind: REF_3D.has(ptg) ? "ref" : "refErr", ixti: f.u16(at + 1), xti: MISSING_XTI, r1: r, r2: r, c1: c, c2: c }
  }
  // PtgArea3d: ptg, ixti u16, first row u32, last row u32, first column u16, last column u16
  if (cce === 15) {
    const ptg = f.u8(at)
    if (!AREA_3D.has(ptg) && !AREA_ERR_3D.has(ptg)) return undefined
    return { kind: AREA_3D.has(ptg) ? "area" : "areaErr", ixti: f.u16(at + 1), xti: MISSING_XTI, r1: f.u32(at + 3), r2: f.u32(at + 7), c1: f.u16(at + 11), c2: f.u16(at + 13) }
  }
  return undefined
}

/** Where XTI `ixti` points (see `Xti`). */
const resolveXti = (books: ReadonlyArray<Book>, xtis: Int32Array, ixti: number): Xti => {
  if (3 * ixti + 2 >= xtis.length) return MISSING_XTI
  const bookIndex = xtis[3 * ixti]!
  const first = xtis[3 * ixti + 1]!
  const last = xtis[3 * ixti + 2]!
  // writers that leave the books out mean this workbook
  const book: Book | undefined = books.length === 0 && bookIndex === 0 ? SELF_BOOK : books[bookIndex]
  if (book === undefined || book.kind === "addin") return NO_BOOK
  if (book.kind === "self") return { kind: "self", first, last }
  // a null id: the book has no link part, so its references are #REF!
  if (book.relationshipId === undefined) return NO_BOOK
  return { kind: "external", book: bookIndex, link: book.link, relationshipId: book.relationshipId, first, last }
}

/** Reads the workbook part. Sheet names are made unique (`uniqueSheetNames`). */
export const readWorkbookBin = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<WorkbookBin, SpreadsheetError> =>
  Effect.suspend(() => {
    const part = entry.name
    const sheets: Array<SheetEntry> = []
    const names: Array<Omit<NameEntry, "ref"> & { readonly ref: RefBuilder | undefined }> = []
    const books: Array<Book> = []
    let externals = 0
    // XTIs of every BrtExternSheet (Excel writes one), appended: a growing array, not a copy per record
    const xtis: Array<number> = []
    let date1904 = false
    const max = limits.maxStringLength
    const reader = new FixedReader()
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      const data = record.data
      // a workbook part starts with BrtBeginBook (its first record is the one at offset 0)
      if (record.offset === 0 && record.type !== BRT_BEGIN_BOOK) return malformed(`not a workbook part: it starts with record 0x${record.type.toString(16)}`, part, record.offset)
      switch (record.type) {
        case BRT_BUNDLE_SH: {
          // hsState u32, iTabID u32, then strRelID (nullable) and strName
          if (sheets.length >= limits.maxSheets) {
            return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: sheets.length + 1 })
          }
          const truncated = "truncated sheet record"
          const f = reader.fixed(data, 8)
          if (f === undefined) return malformed(truncated, part, record.offset)
          const state = f.u32(0)
          const rel = readNullableWideString(data, 8, max)
          if (rel._tag !== "Ok") return stringError(rel, limits, truncated, part, record.offset)
          const name = readWideString(data, rel.end, max)
          if (name._tag !== "Ok") return stringError(name, limits, truncated, part, record.offset)
          sheets.push({
            name: name.value,
            visibility: sheetVisibility(state),
            relationshipId: rel.value
          })
          return undefined
        }
        case BRT_WB_PROP: {
          // flags u32 (f1904 bit 0), ...; a record too short for them, as Excel would, means the 1900 system
          const f = reader.fixed(data, 4)
          date1904 = f !== undefined && (f.u32(0) & 0x01) !== 0
          return undefined
        }
        case BRT_SUP_SELF:
        case BRT_SUP_SAME:
          if (books.length < MAX_BOOKS) books.push(SELF_BOOK)
          return undefined
        case BRT_SUP_ADDIN:
          if (books.length < MAX_BOOKS) books.push(ADDIN_BOOK)
          return undefined
        case BRT_SUP_BOOK_SRC: {
          // numbered whether kept or not, as Excel numbers external links
          const link = ++externals
          if (books.length >= MAX_BOOKS) return undefined
          // a null id is a complete record: the book has no link part, so its references are #REF!
          const rel = readNullableWideString(data, 0, max)
          if (rel._tag !== "Ok") return stringError(rel, limits, "truncated external book record", part, record.offset)
          books.push({ kind: "external", link, relationshipId: rel.value })
          return undefined
        }
        case BRT_EXTERN_SHEET: {
          // cXti u32, then cXti x (iSupBook u32, itabFirst i32, itabLast i32)
          const f = reader.fixed(data, 4)
          // the XTIs are the record's tail: checked against its length as a whole
          if (f === undefined || 4 + 12 * f.u32(0) > data.length) return malformed("truncated external sheet record", part, record.offset)
          const count = f.u32(0)
          // progress: k up to count, which fits in the record (checked above); at most MAX_XTIS kept
          for (let k = 0; k < count && xtis.length < 3 * MAX_XTIS; k++) xtis.push(f.i32(4 + 12 * k), f.i32(8 + 12 * k), f.i32(12 + 12 * k))
          return undefined
        }
        case BRT_NAME: {
          // flags u32 (fHidden bit 0, fBuiltin bit 5), chKey u8, itab u32, name, then the formula: cce u32, rgce
          if (names.length >= limits.maxDefinedNames) {
            return new LimitExceeded({ limit: "maxDefinedNames", max: limits.maxDefinedNames, actual: names.length + 1 })
          }
          // the name's string comes after the fixed fields, so a record it fits in holds them
          const name = readWideString(data, 9, max)
          if (name._tag !== "Ok") return stringError(name, limits, "truncated defined name", part, record.offset)
          // the formula: its length, then that many bytes of tokens, all inside the record
          const f = reader.fixed(data, name.end + 4)
          if (f === undefined || name.end + 4 + f.u32(name.end) > data.length) return malformed("defined name formula past its record", part, record.offset)
          const flags = f.u32(0)
          const itab = f.u32(5)
          const cce = f.u32(name.end)
          const builtin = (flags & 0x20) !== 0
          names.push({
            name: builtin ? `_xlnm.${name.value}` : name.value,
            scope: itab === 0xFFFFFFFF ? undefined : itab,
            hidden: (flags & 0x01) !== 0,
            ref: nameRef(f, name.end + 4, cce)
          })
          return undefined
        }
        default:
          return undefined
      }
    }
    return readRecords(archive.read(entry), handle, limits, part).pipe(
      Effect.filterOrFail((records) => records > 0, () => malformed("empty workbook part", part)),
      Effect.map((): WorkbookBin => {
        const unique = uniqueSheetNames(sheets.map((s) => s.name))
        const xtiTable = Int32Array.from(xtis)
        // each XTI that names use is resolved once (an XTI index is a u16), and shared by them
        const resolved: Array<Xti | undefined> = []
        // progress: one name per iteration (at most maxDefinedNames)
        for (const { ref } of names) {
          if (ref !== undefined) ref.xti = resolved[ref.ixti] ??= resolveXti(books, xtiTable, ref.ixti)
        }
        return { sheets: sheets.map((s, k) => ({ ...s, name: unique[k]! })), date1904, books, xtis: xtiTable, names }
      })
    )
  })

/**
 * The sheet names of an external link part (its BrtSupTabs record), at
 * most `maxSheets`. The part is read up to that record only.
 */
export const readExternalSheets = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<ReadonlyArray<string>, SpreadsheetError> =>
  Effect.suspend(() => {
    const names: Array<string> = []
    const reader = new FixedReader()
    const truncated = "truncated external sheet list"
    const handle = (record: Biff12Record): Handled => {
      if (record.type !== BRT_SUP_TABS) return undefined
      // cTab u32, then cTab strings
      const f = reader.fixed(record.data, 4)
      if (f === undefined) return malformed(truncated, entry.name, record.offset)
      const count = f.u32(0)
      if (count > limits.maxSheets) return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: count })
      let at = 4
      // progress: k up to count (at most maxSheets); each string read advances at
      for (let k = 0; k < count; k++) {
        const name = readWideString(record.data, at, limits.maxStringLength)
        if (name._tag !== "Ok") return stringError(name, limits, truncated, entry.name, record.offset)
        names.push(name.value)
        at = name.end
      }
      // the sheet list is all this part is read for
      return "stop"
    }
    return Effect.as(readRecords(archive.read(entry), handle, limits, entry.name), names)
  })

/**
 * The sheet names of the external books that defined names point into, by
 * book index. Each external link part is read once, however many books
 * point at it (the inflate budget counts a part once, so reading it per
 * book would be free work); a book whose part is missing gets none, so its
 * references are `#REF!`.
 */
export const readExternalBooks = (
  archive: Archive,
  workbook: WorkbookBin,
  partOf: (relationshipId: string) => ZipEntry | undefined,
  limits: Limits
): Effect.Effect<Map<number, ReadonlyArray<string>>, SpreadsheetError> =>
  Effect.gen(function*() {
    const byPart = new Map<ZipEntry, ReadonlyArray<string>>()
    const out = new Map<number, ReadonlyArray<string>>()
    const seen = new Set<number>()
    // progress: one name per iteration (at most maxDefinedNames); each book is looked at once
    for (const n of workbook.names) {
      const xti = n.ref?.xti
      if (xti?.kind !== "external" || seen.has(xti.book)) continue
      seen.add(xti.book)
      const linkEntry = partOf(xti.relationshipId)
      if (linkEntry === undefined) continue
      let names = byPart.get(linkEntry)
      if (names === undefined) {
        names = yield* readExternalSheets(archive, linkEntry, limits)
        byPart.set(linkEntry, names)
      }
      out.set(xti.book, names)
    }
    return out
  })

/**
 * The defined names, their formulas written out. `externalSheets` gives the
 * sheet names of external books (by book index) that could be read; a
 * reference into a book without them, or to a deleted sheet, gives
 * `#REF!$A$1`; one to deleted cells `Sheet!#REF!`; a missing XTI `""`.
 * Names scoped to a sheet that doesn't exist become workbook-wide, as in
 * XLSX.
 */
export const definedNames = (
  workbook: WorkbookBin,
  externalSheets: ReadonlyMap<number, ReadonlyArray<string>>
): Array<DefinedName> => {
  const sheetNames = workbook.sheets.map((s) => s.name)
  const prefix = (xti: Xti): string | undefined => {
    if (xti.kind === "missing") return undefined
    if (xti.kind === "noBook") return "#REF!"
    const names = xti.kind === "self" ? sheetNames : externalSheets.get(xti.book)
    const { first, last } = xti
    if (names === undefined || first < 0 || last < first) return "#REF!"
    const a = names[first]
    const b = names[last]
    if (a === undefined || b === undefined) return "#REF!"
    return sheetPrefix(xti.kind === "self" ? "" : `[${xti.link}]`, first === last ? [a] : [a, b])
  }
  const formula = (ref: NameRef | undefined): string => {
    if (ref === undefined) return ""
    const sheet = prefix(ref.xti)
    if (sheet === undefined) return ""
    switch (ref.kind) {
      case "ref":
        return sheet + cellText(ref.r1, ref.c1)
      case "area":
        return sheet + areaText(ref.r1, ref.r2, ref.c1, ref.c2, BIFF12_GRID)
      default:
        return sheet === "#REF!" ? sheet : `${sheet}#REF!`
    }
  }
  return toDefinedNames(
    // progress: one name per iteration (at most maxDefinedNames)
    workbook.names.map((n) => ({ name: n.name, formula: formula(n.ref), hidden: n.hidden, scope: n.scope })),
    sheetNames.length
  )
}

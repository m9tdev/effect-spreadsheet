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
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { DefinedName, Visibility } from "../../Workbook.ts"
import { Fields, readNullableWideString, readWideString } from "../biff12/fields.ts"
import { type Biff12Record, readRecords } from "../biff12/records.ts"
import { areaText, BIFF12_GRID, cellText, sheetPrefix } from "../refText.ts"
import type { Archive } from "../zip/archive.ts"
import type { ZipEntry } from "../zip/directory.ts"

export const malformed = (reason: string, part?: string, offset?: number) =>
  new MalformedFile({
    format: "xlsb",
    reason,
    ...(part === undefined ? {} : { part }),
    ...(offset === undefined ? {} : { offset })
  })

const tooLong = (limits: Limits) =>
  new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual: limits.maxStringLength + 1 })

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

export interface SheetEntry {
  readonly name: string
  readonly visibility: Visibility
  readonly relationshipId: string | undefined
}

/**
 * A supporting book (BrtSup* records, in order): this workbook, an add-in,
 * or an external workbook, whose relationship leads to its external link
 * part. `link` numbers the external ones from 1, as XLSX does (`[1]`).
 */
export type Book =
  | { readonly kind: "self" }
  | { readonly kind: "addin" }
  | { readonly kind: "external"; readonly link: number; readonly relationshipId: string }

/** A 3-D reference token of a defined name: its XTI and cells. */
export interface NameRef {
  readonly kind: "ref" | "area" | "refErr" | "areaErr"
  readonly ixti: number
  readonly r1: number
  readonly r2: number
  readonly c1: number
  readonly c2: number
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
  readonly names: ReadonlyArray<NameEntry>
}

/** Ptg types of a 3-D reference, area and their deleted-cells forms, in each class (reference, value, array). */
const REF_3D = new Set([0x3A, 0x5A, 0x7A])
const AREA_3D = new Set([0x3B, 0x5B, 0x7B])
const REF_ERR_3D = new Set([0x3C, 0x5C, 0x7C])
const AREA_ERR_3D = new Set([0x3D, 0x5D, 0x7D])

/** The reference of a name formula's tokens (`rgce`, `cce` bytes at `at`), when it is a single 3-D one. */
const nameRef = (f: Fields, at: number, cce: number): NameRef | undefined => {
  const ptg = f.u8(at)
  if (ptg === undefined) return undefined
  const ixti = f.u16(at + 1)!
  // PtgRef3d: ixti, row u32, column u16 (flags in its high bits)
  if (cce === 9 && (REF_3D.has(ptg) || REF_ERR_3D.has(ptg))) {
    const r = f.u32(at + 3)!
    const c = f.u16(at + 7)!
    return { kind: REF_3D.has(ptg) ? "ref" : "refErr", ixti, r1: r, r2: r, c1: c, c2: c }
  }
  // PtgArea3d: ixti, first row u32, last row u32, first column u16, last column u16
  if (cce === 15 && (AREA_3D.has(ptg) || AREA_ERR_3D.has(ptg))) {
    return { kind: AREA_3D.has(ptg) ? "area" : "areaErr", ixti, r1: f.u32(at + 3)!, r2: f.u32(at + 7)!, c1: f.u16(at + 11)!, c2: f.u16(at + 13)! }
  }
  return undefined
}

/** Reads the workbook part. */
export const readWorkbookBin = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<WorkbookBin, SpreadsheetError> =>
  Effect.suspend(() => {
    const part = entry.name
    const sheets: Array<SheetEntry> = []
    const names: Array<NameEntry> = []
    const books: Array<Book> = []
    let externals = 0
    // XTIs of every BrtExternSheet (Excel writes one), appended: a growing array, not a copy per record
    const xtis: Array<number> = []
    let date1904 = false
    const max = limits.maxStringLength
    let records = 0
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      // a workbook part starts with BrtBeginBook
      if (records++ === 0 && record.type !== BRT_BEGIN_BOOK) return malformed(`not a workbook part: it starts with record 0x${record.type.toString(16)}`, part, record.offset)
      const f = new Fields(record.data)
      switch (record.type) {
        case BRT_BUNDLE_SH: {
          // hsState u32, iTabID u32, strRelID (nullable), strName
          if (sheets.length >= limits.maxSheets) {
            return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: sheets.length + 1 })
          }
          const state = f.u32(0)
          const rel = readNullableWideString(record.data, 8, max)
          const name = rel === undefined || rel === null ? rel : readWideString(record.data, rel.end, max)
          if (state === undefined || name === undefined || rel === undefined) return malformed("truncated sheet record", part, record.offset)
          if (name === null || rel === null) return tooLong(limits)
          sheets.push({
            name: name.value,
            visibility: state === 1 ? "hidden" : state === 2 ? "veryHidden" : "visible",
            relationshipId: rel.value
          })
          return undefined
        }
        case BRT_WB_PROP:
          date1904 = ((f.u32(0) ?? 0) & 0x01) !== 0
          return undefined
        case BRT_SUP_SELF:
        case BRT_SUP_SAME:
          books.push({ kind: "self" })
          return undefined
        case BRT_SUP_ADDIN:
          books.push({ kind: "addin" })
          return undefined
        case BRT_SUP_BOOK_SRC: {
          const rel = readWideString(record.data, 0, max)
          if (rel === null) return tooLong(limits)
          books.push({ kind: "external", link: ++externals, relationshipId: rel?.value ?? "" })
          return undefined
        }
        case BRT_EXTERN_SHEET: {
          // cXti u32, then cXti x (iSupBook u32, itabFirst i32, itabLast i32)
          const count = f.u32(0)
          if (count === undefined || 4 + 12 * count > record.data.length) return malformed("truncated external sheet record", part, record.offset)
          // progress: k up to count, which fits in the record (checked above)
          for (let k = 0; k < count; k++) xtis.push(f.u32(4 + 12 * k)! | 0, f.u32(8 + 12 * k)! | 0, f.u32(12 + 12 * k)! | 0)
          return undefined
        }
        case BRT_NAME: {
          // flags u32 (fHidden bit 0, fBuiltin bit 5), chKey u8, itab u32, name, then the formula: cce u32, rgce
          if (names.length >= limits.maxDefinedNames) {
            return new LimitExceeded({ limit: "maxDefinedNames", max: limits.maxDefinedNames, actual: names.length + 1 })
          }
          const flags = f.u32(0)
          const itab = f.u32(5)
          const name = readWideString(record.data, 9, max)
          if (flags === undefined || itab === undefined || name === undefined) return malformed("truncated defined name", part, record.offset)
          if (name === null) return tooLong(limits)
          const cce = f.u32(name.end)
          if (cce === undefined || name.end + 4 + cce > record.data.length) return malformed("defined name formula past its record", part, record.offset)
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
      Effect.filterOrFail(() => records > 0, () => malformed("empty workbook part", part)),
      Effect.map((): WorkbookBin => ({ sheets, date1904, books, xtis: Int32Array.from(xtis), names }))
    )
  })

/** The sheet names of an external link part (its BrtSupTabs record), at most `maxSheets`. */
export const readExternalSheets = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<ReadonlyArray<string>, SpreadsheetError> =>
  Effect.suspend(() => {
    const names: Array<string> = []
    let seen = false
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      if (record.type !== BRT_SUP_TABS || seen) return undefined
      seen = true
      const count = new Fields(record.data).u32(0)
      if (count === undefined) return malformed("truncated external sheet list", entry.name, record.offset)
      if (count > limits.maxSheets) return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: count })
      let at = 4
      // progress: k up to count (at most maxSheets); each string read advances at
      for (let k = 0; k < count; k++) {
        const name = readWideString(record.data, at, limits.maxStringLength)
        if (name === undefined) return malformed("truncated external sheet list", entry.name, record.offset)
        if (name === null) return tooLong(limits)
        names.push(name.value)
        at = name.end
      }
      return undefined
    }
    return Effect.as(readRecords(archive.read(entry), handle, limits, entry.name), names)
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
  const { books, xtis } = workbook
  const prefix = (ixti: number): string | undefined => {
    if (3 * ixti + 2 >= xtis.length) return undefined
    const bookIndex = xtis[3 * ixti]!
    const first = xtis[3 * ixti + 1]!
    const last = xtis[3 * ixti + 2]!
    // writers that leave the books out mean this workbook
    const book: Book | undefined = books.length === 0 && bookIndex === 0 ? { kind: "self" } : books[bookIndex]
    if (book === undefined || book.kind === "addin") return "#REF!"
    const names = book.kind === "self" ? sheetNames : externalSheets.get(bookIndex)
    if (names === undefined || first < 0 || last < first) return "#REF!"
    const a = names[first]
    const b = names[last]
    if (a === undefined || b === undefined) return "#REF!"
    return sheetPrefix(book.kind === "self" ? "" : `[${book.link}]`, first === last ? [a] : [a, b])
  }
  const formula = (ref: NameRef | undefined): string => {
    if (ref === undefined) return ""
    const sheet = prefix(ref.ixti)
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
  // progress: one name per iteration (at most maxDefinedNames)
  return workbook.names.map((n) => ({
    name: n.name,
    formula: formula(n.ref),
    ...(n.scope !== undefined && n.scope < sheetNames.length ? { scope: n.scope } : {}),
    hidden: n.hidden
  }))
}

/** External books that some name's reference points into, by book index. */
export const referencedBooks = (workbook: WorkbookBin): Array<number> => {
  const out = new Set<number>()
  // progress: one name per iteration
  for (const n of workbook.names) {
    if (n.ref === undefined || 3 * n.ref.ixti + 2 >= workbook.xtis.length) continue
    const book = workbook.xtis[3 * n.ref.ixti]!
    if (workbook.books[book]?.kind === "external") out.add(book)
  }
  return [...out]
}

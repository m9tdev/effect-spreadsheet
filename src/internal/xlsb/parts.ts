/**
 * The shared strings (`sharedStrings.bin`) and number formats
 * (`styles.bin`) of an XLSB package.
 *
 * @internal
 */
import { Effect } from "effect"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { Fields, readRichStr, readWideString } from "../biff12/fields.ts"
import { type Biff12Record, readRecords } from "../biff12/records.ts"
import { formatFor, type NumberFormat } from "../xlsx/numFmt.ts"
import type { Archive } from "../zip/archive.ts"
import type { ZipEntry } from "../zip/directory.ts"
import { sliced } from "./slices.ts"
import { malformed, tooLong } from "./workbook.ts"

const BRT_SST_ITEM = 0x0013
const BRT_FMT = 0x002C
const BRT_XF = 0x002F
const BRT_BEGIN_FMTS = 0x0267
const BRT_END_FMTS = 0x0268
const BRT_BEGIN_CELL_XFS = 0x0269
const BRT_END_CELL_XFS = 0x026A

/**
 * Reads the shared-string table: one BrtSSTItem (a RichStr) per string.
 * The counts in BrtBeginSst aren't trusted (nothing is sized by them).
 */
export const readSharedStrings = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<ReadonlyArray<string>, SpreadsheetError> =>
  Effect.suspend(() => {
    const strings: Array<string> = []
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      if (record.type !== BRT_SST_ITEM) return undefined
      if (strings.length >= limits.maxSharedStrings) {
        return new LimitExceeded({ limit: "maxSharedStrings", max: limits.maxSharedStrings, actual: strings.length + 1 })
      }
      const text = readRichStr(record.data, 0, limits.maxStringLength)
      if (text === undefined) return malformed("truncated shared string", entry.name, record.offset)
      if (typeof text === "number") return tooLong(limits, text)
      strings.push(text.value)
      return undefined
    }
    return Effect.as(readRecords(sliced(archive.read(entry)), handle, limits, entry.name), strings)
  })

/**
 * The number formats of the cell styles: a cell's style index (u24) is an
 * index into the BrtXF records of BrtBeginCellXFs. Kept as one u16 format
 * id per style (at most `MAX_CELL_XFS` styles, the u24 range), with one
 * `NumberFormat` per format id, made when a cell first uses it: a flood
 * of tiny BrtXF records costs 2 bytes each.
 */
export interface CellStyles {
  /** The number format of a cell style; `undefined` past the styles. */
  readonly format: (style: number) => NumberFormat | undefined
  /** Styles kept. */
  readonly count: number
}

/** Cell styles kept at most: a cell's style index is 24 bits. */
export const MAX_CELL_XFS = 0x1000000

const cellStyles = (ids: Uint16Array, count: number, custom: ReadonlyMap<number, string>): CellStyles => {
  const byId = new Map<number, NumberFormat>()
  return {
    count,
    format: (style) => {
      if (style >= count) return undefined
      const id = ids[style]!
      let format = byId.get(id)
      if (format === undefined) {
        format = formatFor(id, custom)
        byId.set(id, format)
      }
      return format
    }
  }
}

export const NO_STYLES: CellStyles = cellStyles(new Uint16Array(0), 0, new Map())

/**
 * Reads the number formats (BrtFmt in BrtBeginFmts) and the cell styles
 * that use them (BrtXF in BrtBeginCellXFs), as for XLSX. BrtXF records
 * elsewhere (cell style XFs, differential formats) don't count. Styles
 * past `maxCellXfs` (tests lower it) aren't kept.
 */
export const readStyles = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits,
  maxCellXfs = MAX_CELL_XFS
): Effect.Effect<CellStyles, SpreadsheetError> =>
  Effect.suspend(() => {
    const custom = new Map<number, string>()
    let ids = new Uint16Array(64)
    let count = 0
    let inFmts = false
    let inCellXfs = false
    const handle = (record: Biff12Record): SpreadsheetError | undefined => {
      switch (record.type) {
        case BRT_BEGIN_FMTS:
          inFmts = true
          return undefined
        case BRT_END_FMTS:
          inFmts = false
          return undefined
        case BRT_BEGIN_CELL_XFS:
          inCellXfs = true
          return undefined
        case BRT_END_CELL_XFS:
          inCellXfs = false
          return undefined
        case BRT_FMT: {
          if (!inFmts) return undefined
          // ifmt u16, stFmtCode
          const id = new Fields(record.data).u16(0)
          const code = readWideString(record.data, 2, limits.maxStringLength)
          if (id === undefined || code === undefined) return malformed("truncated number format", entry.name, record.offset)
          if (typeof code === "number") return tooLong(limits, code)
          custom.set(id, code.value)
          return undefined
        }
        case BRT_XF: {
          if (!inCellXfs) return undefined
          // ixfeParent u16, iFmt u16, ...
          const id = new Fields(record.data).u16(2)
          if (id === undefined) return malformed("truncated cell format", entry.name, record.offset)
          // styles past the u24 range can't be used by any cell
          if (count >= maxCellXfs) return undefined
          if (count === ids.length) {
            const grown = new Uint16Array(Math.min(2 * ids.length, maxCellXfs))
            grown.set(ids)
            ids = grown
          }
          ids[count++] = id
          return undefined
        }
        default:
          return undefined
      }
    }
    return Effect.map(readRecords(sliced(archive.read(entry)), handle, limits, entry.name), () => cellStyles(ids, count, custom))
  })

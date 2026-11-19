/**
 * The shared strings (`sharedStrings.bin`) and number formats
 * (`styles.bin`) of an XLSB package.
 *
 * @internal
 */
import { Effect } from "effect"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { readRichStr, readWideString } from "../biff12/fields.ts"
import { type Biff12Record, readRecords } from "../biff12/records.ts"
import { FixedReader } from "../binary/layout.ts"
import { addFormat, cellStyles, type CellStyles, type NumberFormat } from "../numFmt.ts"
import { readers } from "../opc.ts"
import type { Archive, LocatedEntry } from "../zip/archive.ts"
import { malformed, stringError } from "./workbook.ts"

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
  entry: LocatedEntry,
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
      if (text._tag !== "Ok") return stringError(text, limits, "truncated shared string", entry.name, record.offset)
      strings.push(text.value)
      return undefined
    }
    return Effect.as(readRecords(archive.read(entry, readers.open), handle, limits, entry.name), strings)
  })

/** Cell styles kept at most: a cell's style index is 24 bits. */
export const MAX_CELL_XFS = 0x1000000

/**
 * Reads the number formats (BrtFmt in BrtBeginFmts) and the cell styles
 * that use them (BrtXF in BrtBeginCellXFs), as for XLSX. BrtXF records
 * elsewhere (cell style XFs, differential formats) don't count. Kept as one
 * u16 format id per style: a flood of tiny BrtXF records costs 2 bytes
 * each. Styles past `maxCellXfs` (at most `MAX_CELL_XFS`, the u24 range;
 * tests lower it) aren't kept.
 */
export const readStyles = (
  archive: Archive,
  entry: LocatedEntry,
  limits: Limits,
  maxCellXfs = MAX_CELL_XFS
): Effect.Effect<CellStyles, SpreadsheetError> =>
  Effect.suspend(() => {
    const custom = new Map<number, NumberFormat>()
    let ids = new Uint16Array(64)
    let count = 0
    let inFmts = false
    let inCellXfs = false
    const reader = new FixedReader()
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
          const truncated = "truncated number format"
          const f = reader.fixed(record.data, 2)
          if (f === undefined) return malformed(truncated, entry.name, record.offset)
          const id = f.u16(0)
          const code = readWideString(record.data, 2, limits.maxStringLength)
          if (code._tag !== "Ok") return stringError(code, limits, truncated, entry.name, record.offset)
          // classified here, a slice per event-loop turn, not on the cells'
          // turn (the rows stream doesn't yield between the cells of a row);
          // a later BrtFmt with the same id wins
          addFormat(custom, id, code.value)
          return undefined
        }
        case BRT_XF: {
          if (!inCellXfs) return undefined
          // ixfeParent u16, iFmt u16, ...
          const f = reader.fixed(record.data, 4)
          if (f === undefined) return malformed("truncated cell format", entry.name, record.offset)
          const id = f.u16(2)
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
    return Effect.map(readRecords(archive.read(entry, readers.open), handle, limits, entry.name), () => cellStyles(ids, count, custom))
  })

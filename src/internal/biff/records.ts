/**
 * BIFF record reader: a stream of `[type u16][length u16][data]` records,
 * as in the `Workbook` (BIFF8) or `Book` (BIFF5) stream of an `.xls`.
 *
 * Every record advances the offset by `4 + length` (at least 4), and a
 * record that runs past the end of the stream is a `MalformedFile`, so a
 * walk over n bytes ends after at most n / 4 records. Fewer than 4 bytes
 * left at the end (padding) end the walk.
 *
 * `CONTINUE` records (0x003C) carry the rest of a record that didn't fit
 * in 8 224 bytes; `nextWithContinues` gathers them with their record, as
 * separate segments (string data restarts at each segment boundary).
 *
 * @internal
 */
import { LimitExceeded, MalformedFile } from "../../SpreadsheetError.ts"
import type { Segments } from "./segments.ts"

export const CONTINUE = 0x003C
export const BOF_BIFF5_8 = 0x0809
export const EOF = 0x000A

export interface BiffRecord {
  readonly type: number
  /** Offset of the record header in the stream. */
  readonly offset: number
  readonly data: Uint8Array
}

/** A record and the data of the `CONTINUE` records right after it. */
export interface ContinuedRecord extends BiffRecord {
  /**
   * `data` first, then each non-empty `CONTINUE`'s data, as ranges of the
   * stream (8 bytes of storage per segment; empty CONTINUEs are dropped,
   * they continue nothing).
   */
  readonly segments: Segments
}

export class RecordReader {
  /** Offset of the next record header. */
  offset: number
  /** Records read so far: a work counter. */
  records = 0
  private readonly bytes: Uint8Array
  private readonly view: DataView
  private readonly part: string

  constructor(bytes: Uint8Array, part: string, offset = 0) {
    this.bytes = bytes
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    this.part = part
    this.offset = offset
  }

  /** The type of the next record, without reading it; undefined at the end. */
  peekType(): number | undefined {
    return this.offset + 4 <= this.bytes.length ? this.view.getUint16(this.offset, true) : undefined
  }

  /** The next record; undefined at the end of the stream. */
  next(): BiffRecord | undefined | MalformedFile {
    const at = this.offset
    if (at + 4 > this.bytes.length) return undefined
    const type = this.view.getUint16(at, true)
    const length = this.view.getUint16(at + 2, true)
    const end = at + 4 + length
    if (end > this.bytes.length) {
      return new MalformedFile({ format: "xls", part: this.part, offset: at, reason: `record 0x${type.toString(16)} runs past the end of the stream` })
    }
    // progress: offset grows by 4 + length >= 4
    this.offset = end
    this.records++
    return { type, offset: at, data: this.bytes.subarray(at + 4, end) }
  }

  /**
   * The next record with the `CONTINUE` records that follow it. Together,
   * counting each record's 4-byte header (so empty `CONTINUE`s count too),
   * they may not exceed `maxBytes`; past it, a `LimitExceeded` named `limit`
   * (the caller's limit: the SST is capped by a different one than a
   * single string).
   */
  nextWithContinues(maxBytes: number, limit: string): ContinuedRecord | undefined | MalformedFile | LimitExceeded {
    const first = this.next()
    if (first === undefined || first instanceof MalformedFile) return first
    // ranges of the stream, in a growable array: no object per segment
    let bounds = new Uint32Array(16)
    bounds[0] = first.offset + 4
    bounds[1] = first.offset + 4 + first.data.length
    let count = 1
    let total = 4 + first.data.length
    const bytes = this.bytes
    // progress: each iteration consumes one CONTINUE record (offset grows by 4 + length)
    while (this.offset + 4 <= bytes.length && this.view.getUint16(this.offset, true) === CONTINUE) {
      const at = this.offset
      const length = this.view.getUint16(at + 2, true)
      const end = at + 4 + length
      if (end > bytes.length) {
        return new MalformedFile({ format: "xls", part: this.part, offset: at, reason: "record 0x3c runs past the end of the stream" })
      }
      this.offset = end
      this.records++
      total += 4 + length
      if (total > maxBytes) return new LimitExceeded({ limit, max: maxBytes, actual: total })
      if (length === 0) continue // nothing to continue: the offset advanced above
      if (2 * count + 2 > bounds.length) {
        const grown = new Uint32Array(bounds.length * 2)
        grown.set(bounds)
        bounds = grown
      }
      bounds[2 * count] = at + 4
      bounds[2 * count + 1] = end
      count++
    }
    // a view of the used part, no copy (at most 2x the used size is kept)
    return { ...first, segments: { bytes, bounds: bounds.subarray(0, 2 * count), count } }
  }
}

/** A BOF record: the BIFF version and what the substream holds. */
export interface Bof {
  /** 0x0600 for BIFF8, 0x0500 for BIFF5. */
  readonly version: number
  /** 0x0005 workbook globals, 0x0010 worksheet, 0x0020 chart, 0x0040 macro sheet, 0x0006 VB module. */
  readonly kind: number
}

/** Reads a BIFF5/8 BOF record's data; undefined if it is too short. */
export const readBof = (record: BiffRecord): Bof | undefined => {
  if (record.type !== BOF_BIFF5_8 || record.data.length < 4) return undefined
  const view = new DataView(record.data.buffer, record.data.byteOffset, record.data.byteLength)
  return { version: view.getUint16(0, true), kind: view.getUint16(2, true) }
}

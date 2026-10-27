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
  /** `data` first, then each `CONTINUE`'s data. */
  readonly segments: ReadonlyArray<Uint8Array>
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
   * The next record with the `CONTINUE` records that follow it. Their data
   * together may not exceed `maxBytes` (a `LimitExceeded`).
   */
  nextWithContinues(maxBytes: number): ContinuedRecord | undefined | MalformedFile | LimitExceeded {
    const first = this.next()
    if (first === undefined || first instanceof MalformedFile) return first
    const segments = [first.data]
    let total = first.data.length
    // progress: each iteration reads one record (offset advances)
    while (this.peekType() === CONTINUE) {
      const next = this.next()
      if (next === undefined) break
      if (next instanceof MalformedFile) return next
      total += next.data.length
      if (total > maxBytes) return new LimitExceeded({ limit: "maxStringLength", max: maxBytes, actual: total })
      segments.push(next.data)
    }
    return { ...first, segments }
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

/**
 * BIFF record reader: a stream of `[type u16][length u16][data]` records,
 * as in the `Workbook` (BIFF8) or `Book` (BIFF5) stream of an `.xls`.
 *
 * Every record advances the offset by `4 + length` (at least 4), and a
 * record that runs past the end of the stream is a `MalformedFile`, so a
 * walk over n bytes ends after at most n / 4 records. Fewer than 4 bytes
 * left at the end (padding) end the walk.
 *
 * A reader may be given an end before the stream's (a sheet's substream
 * ends where the next one starts): it reads nothing at or past it.
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
  /** Where reading stops: no record may run past it. */
  readonly end: number

  constructor(bytes: Uint8Array, part: string, offset = 0, end = bytes.length) {
    this.bytes = bytes
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    this.part = part
    this.offset = offset
    this.end = Math.min(end, bytes.length)
  }

  private pastEnd(at: number, type: number): MalformedFile {
    const where = this.end === this.bytes.length ? "the stream" : "its substream"
    return new MalformedFile({ format: "xls", part: this.part, offset: at, reason: `record 0x${type.toString(16)} runs past the end of ${where}` })
  }

  /** The type of the next record, without reading it; undefined at the end. */
  peekType(): number | undefined {
    return this.offset + 4 <= this.end ? this.view.getUint16(this.offset, true) : undefined
  }

  /** The next record; undefined at the end of the stream. */
  next(): BiffRecord | undefined | MalformedFile {
    const at = this.offset
    if (at + 4 > this.end) return undefined
    const type = this.view.getUint16(at, true)
    const length = this.view.getUint16(at + 2, true)
    const end = at + 4 + length
    if (end > this.end) return this.pastEnd(at, type)
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
    const gathering = this.startContinues(first, maxBytes, limit)
    const done = this.continues(gathering, Infinity)
    return done === true ? continued(gathering) : done === false ? undefined : done
  }

  /** Starts gathering the `CONTINUE`s after `first` (a record this reader just read), within `maxBytes` (named `limit`). */
  startContinues(first: BiffRecord, maxBytes: number, limit: string): Gathering {
    const bounds = new Uint32Array(16)
    bounds[0] = first.offset + 4
    bounds[1] = first.offset + 4 + first.data.length
    return { first, bytes: this.bytes, maxBytes, limit, bounds, count: 1, total: 4 + first.data.length }
  }

  /**
   * Reads at most `max` more of the `CONTINUE` records after `gathering`'s
   * record (see `nextWithContinues`), so that a long run of them can be
   * read in bounded steps: true once the next record isn't a `CONTINUE`
   * (then `continued` gives the record), false while some may be left.
   */
  continues(gathering: Gathering, max: number): boolean | MalformedFile | LimitExceeded {
    const g = gathering
    // progress: each iteration consumes one CONTINUE record (offset grows by 4 + length), at most max
    for (let k = 0; k < max; k++) {
      if (this.offset + 4 > this.end || this.view.getUint16(this.offset, true) !== CONTINUE) return true
      const at = this.offset
      const length = this.view.getUint16(at + 2, true)
      const end = at + 4 + length
      if (end > this.end) return this.pastEnd(at, CONTINUE)
      this.offset = end
      this.records++
      g.total += 4 + length
      if (g.total > g.maxBytes) return new LimitExceeded({ limit: g.limit, max: g.maxBytes, actual: g.total })
      if (length === 0) continue // nothing to continue: the offset advanced above
      if (2 * g.count + 2 > g.bounds.length) {
        const grown = new Uint32Array(g.bounds.length * 2)
        grown.set(g.bounds)
        g.bounds = grown
      }
      g.bounds[2 * g.count] = at + 4
      g.bounds[2 * g.count + 1] = end
      g.count++
    }
    return this.offset + 4 > this.end || this.view.getUint16(this.offset, true) !== CONTINUE
  }
}

/** A record whose `CONTINUE`s are being read (see `RecordReader.continues`). */
export interface Gathering {
  readonly first: BiffRecord
  /** The reader's stream, which `bounds` index. */
  readonly bytes: Uint8Array
  readonly maxBytes: number
  readonly limit: string
  /** Ranges of the stream, in a growable array: no object per segment. */
  bounds: Uint32Array
  count: number
  total: number
}

/** The record and its `CONTINUE`s gathered so far. */
export const continued = (g: Gathering): ContinuedRecord => {
  // a view of the used part, no copy (at most 2x the used size is kept)
  return { ...g.first, segments: { bytes: g.bytes, bounds: g.bounds.subarray(0, 2 * g.count), count: g.count } }
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

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
 * in 8 224 bytes. A reader checks them (`continues`, or `SubstreamReader`'s
 * continued type) and moves past them; their data is read with the
 * record's through a `RecordCursor` (see `cursor.ts`).
 *
 * `SubstreamReader` reads one substream with the substreams embedded in
 * it (BOF ... EOF, such as a chart object in a sheet) skipped.
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

export class RecordReader {
  /** Offset of the next record header. */
  offset: number
  /** Records read so far: a work counter. */
  records = 0
  protected readonly bytes: Uint8Array
  protected readonly view: DataView
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

  protected pastEnd(at: number, type: number): MalformedFile {
    const where = this.end === this.bytes.length ? "the stream" : "its substream"
    return new MalformedFile({ format: "xls", part: this.part, offset: at, reason: `record 0x${type.toString(16)} runs past the end of ${where}` })
  }

  /** The type of the next record, without reading it; undefined at the end. */
  peekType(): number | undefined {
    return this.offset + 4 <= this.end ? this.view.getUint16(this.offset, true) : undefined
  }

  /**
   * Moves past the record at the offset, its header checked against the
   * end: the offset after it; undefined at the end of the stream, or, with
   * `only`, when the record is of another type (then nothing moves).
   */
  protected advance(only = -1): number | undefined | MalformedFile {
    const at = this.offset
    if (at + 4 > this.end) return undefined
    const type = this.view.getUint16(at, true)
    if (only !== -1 && type !== only) return undefined
    const end = at + 4 + this.view.getUint16(at + 2, true)
    if (end > this.end) return this.pastEnd(at, type)
    // progress: offset grows by 4 + length >= 4
    this.offset = end
    this.records++
    return end
  }

  /** The next record; undefined at the end of the stream. */
  next(): BiffRecord | undefined | MalformedFile {
    const at = this.offset
    const end = this.advance()
    if (typeof end !== "number") return end
    return { type: this.view.getUint16(at, true), offset: at, data: this.bytes.subarray(at + 4, end) }
  }

  /**
   * Reads at most `max` of the `CONTINUE` records after `first`, the record
   * this reader read last (with any of its `CONTINUE`s read by earlier
   * calls), so that a long run of them can be read in bounded steps:
   * "done" once the next record isn't a `CONTINUE` (again on every later
   * call, which reads nothing), "more" while some may be left. The record's
   * data and theirs are then read through a `RecordCursor`. Together,
   * counting each record's 4-byte header (so empty `CONTINUE`s count too),
   * they may not exceed `maxBytes`; past it, a `LimitExceeded` named `limit`
   * (the caller's limit: the SST is capped by a different one than a single
   * string) with their size so far.
   */
  continues(first: BiffRecord, max: number, maxBytes: number, limit: string): "done" | "more" | MalformedFile | LimitExceeded {
    // the headers are read here rather than through `advance`: a run of
    // empty CONTINUEs is a loop over headers only, and the call costs about
    // a fifth of it. progress: each iteration moves past one CONTINUE
    // (the offset grows by 4 + length >= 4), at most max
    for (let k = 0; k < max; k++) {
      const at = this.offset
      if (at + 4 > this.end || this.view.getUint16(at, true) !== CONTINUE) return "done"
      const end = at + 4 + this.view.getUint16(at + 2, true)
      if (end > this.end) return this.pastEnd(at, CONTINUE)
      this.offset = end
      this.records++
      // the record and its CONTINUEs lie end to end: their size is the distance
      const total = end - first.offset
      if (total > maxBytes) return new LimitExceeded({ limit, max: maxBytes, actual: total })
    }
    return this.peekType() === CONTINUE ? "more" : "done"
  }
}

/**
 * A record of a substream embedded in the one being read (see
 * `SubstreamReader`), skipped: the walkers treat every one alike. The first
 * of them (`opens`) is the embedded substream's BOF; the others come one
 * per call only so that each call reads one record, as `RecordReader.next`
 * does.
 */
export class Embedded {
  /** Offset of the record header in the stream. */
  readonly offset: number
  /** Whether this is the BOF that opens the embedded substream. */
  readonly opens: boolean

  constructor(offset: number, opens: boolean) {
    this.offset = offset
    this.opens = opens
  }
}

/**
 * The records of one substream (a sheet's, after its BOF), with every
 * substream embedded in it skipped: an embedded chart sits between a BOF
 * and its EOF inside the sheet, and may hold substreams of its own. Their
 * records come from `read` as `Embedded`, so a walker never reads them;
 * the substream's own EOF comes as a record like any other. One record per
 * call, as `next`; errors and the end are the same as `next`'s. (`next`
 * itself reads records as they are, for the substream's own BOF.)
 *
 * A record of the substream of type `continued` (a BIFF8 STRING, for the
 * rows) comes with its `CONTINUE` records: the call moves past them too,
 * checking each, and the record's data and theirs are read through a
 * `RecordCursor`. Other records' `CONTINUE`s come as records.
 */
export class SubstreamReader extends RecordReader {
  /** Embedded substreams open at the reader's offset (each one opened by a BOF read: bounded by the records read). */
  private depth = 0
  /** The type whose `CONTINUE`s are read with it; -1 for none. */
  private readonly continued: number

  constructor(bytes: Uint8Array, part: string, offset = 0, end = bytes.length, continued = -1) {
    super(bytes, part, offset, end)
    this.continued = continued
  }

  /** The next record of this substream, or an embedded substream's; undefined at the end of the stream. */
  read(): BiffRecord | Embedded | undefined | MalformedFile {
    // the header is read here rather than through `next`: this is the
    // sheet walkers' loop over every record, and the extra call costs
    // about a tenth of a bare record loop
    const at = this.offset
    if (at + 4 > this.end) return undefined
    const type = this.view.getUint16(at, true)
    const end = at + 4 + this.view.getUint16(at + 2, true)
    if (end > this.end) return this.pastEnd(at, type)
    // progress: offset grows by 4 + length >= 4
    this.offset = end
    this.records++
    if (this.depth > 0) {
      if (type === BOF_BIFF5_8) this.depth++
      else if (type === EOF) this.depth--
      return new Embedded(at, false)
    }
    if (type === BOF_BIFF5_8) {
      this.depth++
      return new Embedded(at, true)
    }
    const record = { type, offset: at, data: this.bytes.subarray(at + 4, end) }
    if (type === this.continued) {
      // progress: each iteration moves past one CONTINUE (advance moves the offset), or returns
      for (let next = this.advance(CONTINUE); next !== undefined; next = this.advance(CONTINUE)) {
        if (next instanceof MalformedFile) return next
      }
    }
    return record
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

/**
 * BIFF12 record reader: the record stream of an XLSB part
 * (`xl/workbook.bin`, a worksheet, `sharedStrings.bin`, `styles.bin`).
 *
 * A record is `[type][size][data]`. The type is a variable-length integer
 * of 1 or 2 bytes, the size one of 1 to 4 bytes: 7 bits per byte, low bits
 * first, the high bit set on every byte but the last. Longer forms are a
 * `MalformedFile` (so no header is over 6 bytes), and a size over
 * `maxRecordBytes` is a `LimitExceeded` before any of the data is held.
 *
 * Parts are fed chunk by chunk, as they inflate (`RecordSplitter.feed`):
 * a record inside one chunk is a view of it; one that spans chunks is
 * copied once into a buffer of its size. Every byte of input is consumed
 * by exactly one header or record body, so a part of n bytes gives at most
 * n / 2 records (the smallest record, `[type][0]`, takes 2 bytes), and the
 * work is linear in n. A record cut off by the end of the part is a
 * `MalformedFile` (`end`).
 *
 * @internal
 */
import { Effect, Stream } from "effect"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"

export interface Biff12Record {
  readonly type: number
  /** Offset of the record header in the part. */
  readonly offset: number
  readonly data: Uint8Array
}

/** What a record handler gives back: nothing, or the error that stops the read. */
export type RecordHandler = (record: Biff12Record) => SpreadsheetError | undefined

const EMPTY = new Uint8Array(0)

/** Where the splitter is: in a header (type, then size), or in a record body. */
const TYPE = 0
const SIZE = 1
const BODY = 2

export class RecordSplitter {
  /** Offset in the part of the next byte to consume. */
  offset = 0
  /** Records completed so far: a work counter. */
  records = 0
  private readonly part: string
  private readonly maxRecordBytes: number
  private state = TYPE
  /** Offset of the current record's header. */
  private start = 0
  /** Header bytes of the current varint read so far. */
  private varintBytes = 0
  private type = 0
  private size = 0
  /** The body of a record that spans chunks, and how much of it is filled. */
  private body: Uint8Array | undefined
  private filled = 0

  constructor(part: string, maxRecordBytes: number) {
    this.part = part
    this.maxRecordBytes = maxRecordBytes
  }

  private malformed(reason: string): MalformedFile {
    return new MalformedFile({ format: "xlsb", part: this.part, offset: this.start, reason })
  }

  private emit(onRecord: RecordHandler, data: Uint8Array): SpreadsheetError | undefined {
    this.records++
    this.state = TYPE
    this.varintBytes = 0
    this.body = undefined
    return onRecord({ type: this.type, offset: this.start, data })
  }

  /**
   * Consumes `chunk`, calling `onRecord` for every record it completes, in
   * order. Stops at, and returns, the first error (from the input or the
   * handler); the splitter isn't usable after one.
   */
  feed(chunk: Uint8Array, onRecord: RecordHandler): SpreadsheetError | undefined {
    let at = 0
    const n = chunk.length
    // progress: every iteration consumes at least one byte of the chunk
    // (at grows: a header byte, or body bytes, take >= 1 while the body
    // isn't full)
    while (at < n) {
      if (this.state === TYPE) {
        if (this.varintBytes === 0) {
          this.start = this.offset + at
          this.type = 0
        }
        const byte = chunk[at++]!
        this.type |= (byte & 0x7F) << (7 * this.varintBytes)
        this.varintBytes++
        if ((byte & 0x80) === 0) {
          this.state = SIZE
          this.varintBytes = 0
          this.size = 0
        } else if (this.varintBytes === 2) {
          this.offset += at
          return this.malformed("record type longer than 2 bytes")
        }
      } else if (this.state === SIZE) {
        const byte = chunk[at++]!
        // 4 x 7 bits: at most 2^28 - 1, no overflow
        this.size += (byte & 0x7F) * 2 ** (7 * this.varintBytes)
        this.varintBytes++
        if ((byte & 0x80) !== 0) {
          if (this.varintBytes === 4) {
            this.offset += at
            return this.malformed(`record 0x${this.type.toString(16)}: size longer than 4 bytes`)
          }
          continue
        }
        if (this.size > this.maxRecordBytes) {
          this.offset += at
          return new LimitExceeded({ limit: "maxRecordBytes", max: this.maxRecordBytes, actual: this.size })
        }
        if (this.size === 0) {
          const error = this.emit(onRecord, EMPTY)
          if (error !== undefined) {
            this.offset += at
            return error
          }
        } else if (n - at >= this.size) {
          // the whole body is in this chunk: a view of it
          const data = chunk.subarray(at, at + this.size)
          at += this.size
          const error = this.emit(onRecord, data)
          if (error !== undefined) {
            this.offset += at
            return error
          }
        } else {
          // spans chunks: one buffer of the record's size (at most maxRecordBytes), filled as they come
          this.state = BODY
          this.body = new Uint8Array(this.size)
          this.filled = 0
        }
      } else {
        const body = this.body!
        const take = Math.min(n - at, body.length - this.filled)
        body.set(chunk.subarray(at, at + take), this.filled)
        at += take
        this.filled += take
        if (this.filled === body.length) {
          const error = this.emit(onRecord, body)
          if (error !== undefined) {
            this.offset += at
            return error
          }
        }
      }
    }
    this.offset += n
    return undefined
  }

  /** The end of the part: a `MalformedFile` if a record is cut off. */
  end(): MalformedFile | undefined {
    if (this.state === TYPE && this.varintBytes === 0) return undefined
    const what = this.state === BODY
      ? `record 0x${this.type.toString(16)} runs past the end of the part (${this.filled} of ${this.size} bytes)`
      : "record header cut off by the end of the part"
    return this.malformed(what)
  }
}

/** The result of `readRecords`: the records read, for cost checks. */
export interface RecordStats {
  readonly records: number
}

/**
 * Reads a part's records from its byte stream, calling `onRecord` for each.
 * Fails with the first error (from the stream, the record layer or the
 * handler). Each chunk is at most one inflate piece, so the work between
 * the stream's turns is bounded.
 */
export const readRecords = <E, R>(
  bytes: Stream.Stream<Uint8Array, E, R>,
  onRecord: RecordHandler,
  limits: Pick<Limits, "maxRecordBytes">,
  part: string
): Effect.Effect<RecordStats, E | SpreadsheetError, R> =>
  Effect.suspend(() => {
    const splitter = new RecordSplitter(part, limits.maxRecordBytes)
    return Stream.runForEach(bytes, (chunk) => {
      const error = splitter.feed(chunk, onRecord)
      return error === undefined ? Effect.void : Effect.fail(error)
    }).pipe(
      Effect.flatMap(() => {
        const error = splitter.end()
        return error === undefined ? Effect.succeed({ records: splitter.records }) : Effect.fail(error)
      })
    )
  })

/** Every record of a whole part (for tests and small parts). */
export const splitAll = (bytes: Uint8Array, part: string, maxRecordBytes: number): Array<Biff12Record> | SpreadsheetError => {
  const out: Array<Biff12Record> = []
  const splitter = new RecordSplitter(part, maxRecordBytes)
  const error = splitter.feed(bytes, (record) => {
    out.push(record)
    return undefined
  }) ?? splitter.end()
  return error ?? out
}

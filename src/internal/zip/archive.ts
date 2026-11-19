/**
 * A zip archive opened for reading: the central directory checked against
 * the local headers, and entries read as byte streams with every limit
 * enforced while they inflate.
 *
 * - Every entry's local header must match its directory entry (signature,
 *   name) and its data must lie inside the buffer; entries may not overlap
 *   or repeat a name (names compare case-insensitively, as OPC parts do).
 * - Inflated bytes count against the entry's declared size and against one
 *   budget for the whole archive, `inflateBudget`: max(`inflateFloorBytes`,
 *   `maxInflateRatio` × archive size), capped at `maxDecompressedBytes`.
 *   Each read names its reader (a sheet, the open-time reads, ...): an
 *   entry counts, per reader, up to the most that reader has inflated of
 *   it. So a reader reading an entry again doesn't spend the budget again
 *   (a sheet can be streamed twice, even after another entry went over
 *   the budget), while two readers pointed at one entry each pay for what
 *   they read: aliasing an entry costs what it reads, never nothing. Many
 *   entries, however small each one is, can't add up past the budget. A
 *   read may be given other limits than the archive's: its budget comes
 *   from them, checked against the same archive-wide total.
 *   The CRC-32 is checked at the end. Decompression only runs as far as
 *   the stream is read.
 * - A read emits one slice of at most `SLICE_BYTES` per pull, after a
 *   yield to the event loop, whatever size the decompressor's output
 *   comes in: the readers of an entry's bytes do no slicing or yielding of
 *   their own. Limits and the CRC-32 are applied to each decompressed (or,
 *   for stored entries, `SLICE_BYTES`) piece as it is taken, before its
 *   slices are emitted.
 *
 * @internal
 */
import { Cause, Effect, Scope, Stream } from "effect"
import * as Channel from "effect/Channel"
import { absurd } from "effect/Function"
import * as Result from "effect/Result"
import type * as Arr from "effect/Array"
import type * as Pull from "effect/Pull"
import { Encrypted, LimitExceeded, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { SLICE_BYTES } from "../bytes.ts"
import { yieldNow } from "../yield.ts"
import { WorkMeter } from "../meter.ts"
import { CRC32_START, crc32Final, crc32Update } from "./crc32.ts"
import { type Directory, normalizeName, readDirectory, type ZipEntry } from "./directory.ts"
import { type InflateRaw, inflateRawWeb } from "./inflate.ts"

/** The archive-wide inflate budget for an archive of `archiveBytes` bytes. */
export const inflateBudget = (archiveBytes: number, limits: Limits): number =>
  Math.min(limits.maxDecompressedBytes, Math.max(limits.inflateFloorBytes, limits.maxInflateRatio * archiveBytes))

const LOCAL = 0x04034b50
const STORED = 0
const DEFLATE = 8

/**
 * How an entry's data can be read, decided once when the archive opens:
 * stored or deflated data, or the reason a read fails. When several apply,
 * encryption comes first, then an unknown compression method, then a stored
 * entry whose two sizes differ.
 */
export type EntryKind =
  | "stored"
  | "deflate"
  | "encrypted"
  | { readonly unsupported: number }
  | "bad-stored-size"

/** A directory entry as an opened archive has checked it: where its data lies in the archive, and how it reads. */
export interface LocatedEntry extends ZipEntry {
  /** The entry's compressed data, [start, end) in the archive's bytes. */
  readonly data: { readonly start: number; readonly end: number }
  readonly kind: EntryKind
}

/**
 * Who reads an entry, for the inflate budget: re-reads under one key are
 * free up to what that key has already inflated of the entry; reads under
 * different keys are charged separately. One key per sheet (`sheet:3`),
 * one for the reads made while a workbook opens (`open`), and so on.
 */
export type ReaderKey = string

export interface Archive {
  readonly entries: ReadonlyArray<LocatedEntry>
  /** The entry with this name (case-insensitive, a leading `/` ignored), if any. */
  readonly find: (name: string) => LocatedEntry | undefined
  /**
   * The entry's uncompressed bytes, one slice of at most `SLICE_BYTES`
   * per pull, each after a yield. Re-runnable. `reader` is who reads it
   * (see `ReaderKey`); `limits` (default: the archive's) set this read's
   * inflate budget.
   */
  readonly read: (entry: LocatedEntry, reader: ReaderKey, limits?: Limits) => Stream.Stream<Uint8Array, SpreadsheetError>
}

const malformed = (reason: string, part?: string, offset?: number) =>
  new MalformedFile({
    format: "zip",
    reason,
    ...(part === undefined ? {} : { part }),
    ...(offset === undefined ? {} : { offset })
  })

/** A failed result with a `MalformedFile`. */
const fail = (reason: string, part?: string, offset?: number) => Result.fail(malformed(reason, part, offset))

/** How `entry` reads (see `EntryKind`). */
const kindOf = (entry: ZipEntry): EntryKind => {
  if ((entry.flags & 0x1) !== 0) return "encrypted"
  if (entry.method === STORED) return entry.compressedSize === entry.uncompressedSize ? "stored" : "bad-stored-size"
  if (entry.method === DEFLATE) return "deflate"
  return { unsupported: entry.method }
}

/** Decodes local header names, as the directory decodes its own (`readDirectory`). */
const names = new TextDecoder("utf-8")

/**
 * Whether the `length` name bytes at `start` decode (as UTF-8) to `name`.
 * ASCII bytes equal to `name`'s characters are compared directly, without
 * decoding; anything else is decoded.
 */
const sameName = (bytes: Uint8Array, start: number, length: number, name: string): boolean => {
  if (length === name.length) {
    let i = 0
    // progress: i increases up to length
    while (i < length && bytes[start + i]! < 0x80 && bytes[start + i] === name.charCodeAt(i)) i++
    if (i === length) return true
  }
  return names.decode(bytes.subarray(start, start + length)) === name
}

/**
 * Checks one entry's local header against its directory entry (signature,
 * compression method, name) and that its data lies in the buffer, before
 * the central directory; the entry with where its data is, and how it
 * reads.
 */
export const locateEntry = (
  bytes: Uint8Array,
  directory: Directory,
  entry: ZipEntry
): Result.Result<LocatedEntry, MalformedFile> =>
  locateIn(bytes, new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), directory, entry)

const locateIn = (
  bytes: Uint8Array,
  view: DataView,
  directory: Directory,
  entry: ZipEntry
): Result.Result<LocatedEntry, MalformedFile> => {
  const at = entry.localOffset
  if (at + 30 > bytes.length) return fail("local header out of bounds", entry.name, at)
  if (view.getUint32(at, true) !== LOCAL) return fail("bad local header signature", entry.name, at)
  if (view.getUint16(at + 8, true) !== entry.method) {
    return fail("local header compression method differs from the directory", entry.name, at)
  }
  const nameLength = view.getUint16(at + 26, true)
  const extraLength = view.getUint16(at + 28, true)
  const start = at + 30 + nameLength + extraLength
  const end = start + entry.compressedSize
  if (end > bytes.length) return fail("entry data out of bounds", entry.name, at)
  // entry data comes before the central directory (and so the EOCD)
  if (end > directory.directoryStart) return fail("entry data overlaps the central directory", entry.name, at)
  // The local name must be the directory's name (decoded the same way); a
  // mismatch means the directory points at the wrong data.
  if (!sameName(bytes, at + 30, nameLength, entry.name)) {
    return fail("local header name differs from the directory", entry.name, at)
  }
  return Result.succeed({ ...entry, data: { start, end }, kind: kindOf(entry) })
}

/** Checks the local headers and the layout; the entries, in directory order, with where each one's data is. */
const locate = (bytes: Uint8Array, directory: Directory): Result.Result<ReadonlyArray<LocatedEntry>, MalformedFile> => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const located: Array<LocatedEntry> = []
  // progress: one entry per iteration (entries are bounded by maxZipEntries)
  for (const entry of directory.entries) {
    const one = locateIn(bytes, view, directory, entry)
    if (Result.isFailure(one)) return Result.fail(one.failure)
    located.push(one.success)
  }
  // No two entries may share bytes: sort by header offset and check the gaps.
  const ordered = [...located].sort((a, b) => a.localOffset - b.localOffset)
  // progress: i increases up to ordered.length
  for (let i = 1; i < ordered.length; i++) {
    const before = ordered[i - 1]!
    const after = ordered[i]!
    if (before.data.end > after.localOffset) {
      return fail("entries overlap", after.name, after.localOffset)
    }
  }
  return Result.succeed(located)
}

/** Opens an archive. `inflate` is replaceable for other inflate implementations (and tests). */
export const open = (
  bytes: Uint8Array,
  limits: Limits,
  inflate: InflateRaw = inflateRawWeb
): Result.Result<Archive, MalformedFile | LimitExceeded> => {
  const read = readDirectory(bytes, limits)
  return Result.isFailure(read) ? Result.fail(read.failure) : openDirectory(bytes, read.success, limits, inflate)
}

/** Opens an archive whose central directory `readDirectory` has read. */
export const openDirectory = (
  bytes: Uint8Array,
  directory: Directory,
  limits: Limits,
  inflate: InflateRaw = inflateRawWeb
): Result.Result<Archive, MalformedFile> => {
  // Names are checked before the local headers.
  const keys: Array<string> = []
  const seen = new Set<string>()
  // progress: one entry per iteration
  for (const entry of directory.entries) {
    const key = normalizeName(entry.name)
    if (seen.has(key)) return fail("duplicate entry name", entry.name, entry.localOffset)
    seen.add(key)
    keys.push(key)
  }
  const located = locate(bytes, directory)
  if (Result.isFailure(located)) return Result.fail(located.failure)
  const entries = located.success
  const byName = new Map<string, LocatedEntry>()
  // progress: i increases up to entries.length (one key per entry, in the same order)
  for (let i = 0; i < entries.length; i++) byName.set(keys[i]!, entries[i]!)
  return Result.succeed(archive(bytes, limits, inflate, entries, byName))
}

/** The archive over entries `open` has checked and located. */
const archive = (
  bytes: Uint8Array,
  limits: Limits,
  inflate: InflateRaw,
  entries: ReadonlyArray<LocatedEntry>,
  byName: ReadonlyMap<string, LocatedEntry>
): Archive => {
  /** Bytes counted against the budget: per reader and entry, the most that reader has inflated of it so far. */
  let charged = 0
  /** Per reader key, each entry's high-water mark. */
  const highWater = new Map<ReaderKey, Map<LocatedEntry, number>>()
  /** The high-water marks of `reader`'s entries. */
  const marksOf = (reader: ReaderKey): Map<LocatedEntry, number> => {
    let marks = highWater.get(reader)
    if (marks === undefined) {
      marks = new Map()
      highWater.set(reader, marks)
    }
    return marks
  }

  /**
   * Counts output bytes of `entry` (`soFar` of them so far) under
   * `limits`, for the reader whose marks are `marks`; an error if a limit
   * is crossed. Only bytes past that reader's high-water mark on the entry
   * are charged, so its re-reads of what it already inflated pass. A
   * chunk that would take the total over the budget is rejected without
   * being charged or raising the high-water mark, so a re-read fails at
   * the same point again.
   */
  const account = (
    entry: LocatedEntry,
    marks: Map<LocatedEntry, number>,
    soFar: number,
    limits: Limits
  ): SpreadsheetError | undefined => {
    if (soFar > entry.uncompressedSize) {
      return malformed(`entry inflates past its declared size of ${entry.uncompressedSize} bytes`, entry.name)
    }
    const before = marks.get(entry) ?? 0
    if (soFar <= before) return undefined
    const total = charged + (soFar - before)
    const budget = inflateBudget(bytes.length, limits)
    if (total > budget) {
      return budget < limits.maxDecompressedBytes
        ? new LimitExceeded({ limit: "maxInflateRatio", max: limits.maxInflateRatio, actual: Math.ceil(total / Math.max(1, bytes.length)) })
        : new LimitExceeded({ limit: "maxDecompressedBytes", max: limits.maxDecompressedBytes, actual: total })
    }
    charged = total
    marks.set(entry, soFar)
    return undefined
  }

  const finish = (entry: LocatedEntry, size: number, crc: number): SpreadsheetError | undefined => {
    if (size !== entry.uncompressedSize) {
      return malformed(`entry is ${size} bytes, its directory says ${entry.uncompressedSize}`, entry.name)
    }
    if (crc !== entry.crc32) return malformed("CRC-32 mismatch", entry.name)
    return undefined
  }

  const read = (entry: LocatedEntry, reader: ReaderKey, readLimits: Limits = limits): Stream.Stream<Uint8Array, SpreadsheetError> => {
    const kind = entry.kind
    if (typeof kind === "object") {
      return Stream.fail(new UnsupportedFormat({ format: "zip", reason: `compression method ${kind.unsupported} in ${entry.name}` }))
    }
    switch (kind) {
      case "encrypted":
        return Stream.fail(new Encrypted({ format: "zip" }))
      case "bad-stored-size":
        return Stream.fail(malformed("stored entry with different compressed and uncompressed sizes", entry.name))
      case "stored":
      case "deflate":
        return readData(entry, kind, reader, readLimits)
      default:
        return absurd(kind)
    }
  }

  /** The bytes of a readable entry, stored or deflated. */
  const readData = (
    entry: LocatedEntry,
    kind: "stored" | "deflate",
    reader: ReaderKey,
    readLimits: Limits
  ): Stream.Stream<Uint8Array, SpreadsheetError> => {
    const data = bytes.subarray(entry.data.start, entry.data.end)
    const marks = marksOf(reader)
    return Stream.fromChannel(Channel.fromTransform((_, scope: Scope.Scope) =>
      Effect.gen(function*() {
        let size = 0
        let crc = CRC32_START
        let done = false
        const inflater = kind === "deflate" ? inflate(data) : undefined
        if (inflater !== undefined) yield* Scope.addFinalizer(scope, Effect.sync(() => inflater.cancel()))
        let storedAt = 0
        // the work meter counts every output byte, re-reads included
        const meter = yield* WorkMeter
        /** The piece being handed out, and how far. */
        let piece: Uint8Array = new Uint8Array(0)
        let pieceAt = 0
        const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Uint8Array>, SpreadsheetError> = Effect.gen(function*() {
          // progress: each iteration emits a non-empty slice of the piece
          // (pieceAt grows), ends, or takes one more piece: a stored slice
          // (storedAt grows) or a chunk from the decompressor (which
          // consumes input)
          while (true) {
            if (pieceAt < piece.length) {
              yield* yieldNow
              const slice = piece.subarray(pieceAt, Math.min(piece.length, pieceAt + SLICE_BYTES))
              pieceAt += slice.length
              return [slice] as const
            }
            if (done) return yield* Cause.done()
            let chunk: Uint8Array | undefined
            if (inflater === undefined) {
              if (storedAt < data.length) {
                chunk = data.subarray(storedAt, Math.min(data.length, storedAt + SLICE_BYTES))
                storedAt += chunk.length
              }
            } else {
              chunk = yield* Effect.tryPromise({
                try: () => inflater.next(),
                catch: (error) =>
                  malformed(`invalid deflate data: ${error instanceof Error && error.message !== "" ? error.message : "corrupt or truncated"}`, entry.name)
              })
            }
            if (chunk === undefined) {
              done = true
              inflater?.cancel()
              const error = finish(entry, size, crc32Final(crc))
              if (error !== undefined) return yield* Effect.fail(error)
              continue // done: the next iteration ends the stream
            }
            if (chunk.length === 0) continue // the decompressor advanced its input
            size += chunk.length
            meter.add("bytes", chunk.length)
            const error = account(entry, marks, size, readLimits)
            if (error !== undefined) {
              done = true
              inflater?.cancel()
              return yield* Effect.fail(error)
            }
            crc = crc32Update(crc, chunk)
            piece = chunk
            pieceAt = 0
          }
        })
        return pull
      })
    ))
  }

  return {
    entries,
    find: (name) => byName.get(normalizeName(name)),
    read
  }
}


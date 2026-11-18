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
 *   Reading an entry again doesn't spend the budget again (each entry
 *   counts up to the most it has inflated), so a sheet can be streamed
 *   twice, even after another entry went over the budget; many entries,
 *   however small each one is, can't add up past it. A read may be given
 *   other limits than the archive's: its budget comes from them, checked
 *   against the same archive-wide total.
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

export interface Archive {
  readonly entries: ReadonlyArray<LocatedEntry>
  /** The entry with this name (case-insensitive, a leading `/` ignored), if any. */
  readonly find: (name: string) => LocatedEntry | undefined
  /**
   * The entry's uncompressed bytes, one slice of at most `SLICE_BYTES`
   * per pull, each after a yield. Re-runnable. `limits`
   * (default: the archive's) set this read's inflate budget.
   */
  readonly read: (entry: LocatedEntry, limits?: Limits) => Stream.Stream<Uint8Array, SpreadsheetError>
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

/** Checks the local headers and the layout; the entries, in directory order, with where each one's data is. */
const locate = (bytes: Uint8Array, directory: Directory): Result.Result<ReadonlyArray<LocatedEntry>, MalformedFile> => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const names = new TextDecoder("utf-8")
  const located: Array<LocatedEntry> = []
  // progress: one entry per iteration (entries are bounded by maxZipEntries)
  for (const entry of directory.entries) {
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
    if (names.decode(bytes.subarray(at + 30, at + 30 + nameLength)) !== entry.name) {
      return fail("local header name differs from the directory", entry.name, at)
    }
    located.push({ ...entry, data: { start, end }, kind: kindOf(entry) })
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
  if (Result.isFailure(read)) return Result.fail(read.failure)
  const directory = read.success
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
  /** Bytes counted against the budget: per entry, the most it has inflated so far. */
  let charged = 0
  const highWater = new Map<LocatedEntry, number>()

  /**
   * Counts output bytes of `entry` (`soFar` of them so far) under
   * `limits`; an error if a limit is crossed. Only bytes past the entry's
   * high-water mark are charged, so re-reads of what was already inflated
   * pass. A chunk that would take the total over the budget is rejected
   * without being charged or raising the high-water mark, so a re-read
   * fails at the same point again.
   */
  const account = (entry: LocatedEntry, soFar: number, limits: Limits): SpreadsheetError | undefined => {
    if (soFar > entry.uncompressedSize) {
      return malformed(`entry inflates past its declared size of ${entry.uncompressedSize} bytes`, entry.name)
    }
    const before = highWater.get(entry) ?? 0
    if (soFar <= before) return undefined
    const total = charged + (soFar - before)
    const budget = inflateBudget(bytes.length, limits)
    if (total > budget) {
      return budget < limits.maxDecompressedBytes
        ? new LimitExceeded({ limit: "maxInflateRatio", max: limits.maxInflateRatio, actual: Math.ceil(total / Math.max(1, bytes.length)) })
        : new LimitExceeded({ limit: "maxDecompressedBytes", max: limits.maxDecompressedBytes, actual: total })
    }
    charged = total
    highWater.set(entry, soFar)
    return undefined
  }

  const finish = (entry: LocatedEntry, size: number, crc: number): SpreadsheetError | undefined => {
    if (size !== entry.uncompressedSize) {
      return malformed(`entry is ${size} bytes, its directory says ${entry.uncompressedSize}`, entry.name)
    }
    if (crc !== entry.crc32) return malformed("CRC-32 mismatch", entry.name)
    return undefined
  }

  const read = (entry: LocatedEntry, readLimits: Limits = limits): Stream.Stream<Uint8Array, SpreadsheetError> => {
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
        return readData(entry, kind, readLimits)
      default:
        return absurd(kind)
    }
  }

  /** The bytes of a readable entry, stored or deflated. */
  const readData = (
    entry: LocatedEntry,
    kind: "stored" | "deflate",
    readLimits: Limits
  ): Stream.Stream<Uint8Array, SpreadsheetError> => {
    const data = bytes.subarray(entry.data.start, entry.data.end)
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
              yield* Effect.yieldNow
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
            const error = account(entry, size, readLimits)
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


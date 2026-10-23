/**
 * A zip archive opened for reading: the central directory checked against
 * the local headers, and entries read as byte streams with every limit
 * enforced while they inflate.
 *
 * - Every entry's local header must match its directory entry (signature,
 *   name) and its data must lie inside the buffer; entries may not overlap
 *   or repeat a name (names compare case-insensitively, as OPC parts do).
 * - Inflated bytes count against `maxDecompressedBytes` for the whole
 *   archive and against the entry's declared size; once an entry has
 *   inflated past `RATIO_GRACE_BYTES`, its ratio to the compressed size
 *   must stay within `maxCompressionRatio`. The CRC-32 is checked at the
 *   end. Decompression only runs as far as the stream is read.
 *
 * @internal
 */
import { Cause, Effect, Scope, Stream } from "effect"
import * as Channel from "effect/Channel"
import type * as Arr from "effect/Array"
import type * as Pull from "effect/Pull"
import { Encrypted, LimitExceeded, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { CRC32_START, crc32Final, crc32Update } from "./crc32.ts"
import { type Directory, readDirectory, type ZipEntry } from "./directory.ts"
import { type InflateRaw, inflateRawWeb } from "./inflate.ts"

/** The ratio limit applies once an entry has inflated past this many bytes. */
export const RATIO_GRACE_BYTES = 1024 * 1024

/** Stored entries are emitted in slices of this size. */
const STORED_SLICE = 64 * 1024

const LOCAL = 0x04034b50
const STORED = 0
const DEFLATE = 8

export interface Archive {
  readonly entries: ReadonlyArray<ZipEntry>
  /** The entry with this name (case-insensitive, a leading `/` ignored), if any. */
  readonly find: (name: string) => ZipEntry | undefined
  /** The entry's uncompressed bytes, chunk by chunk. Re-runnable. */
  readonly read: (entry: ZipEntry) => Stream.Stream<Uint8Array, SpreadsheetError>
  /** The whole entry; fails up front if its declared size is over `maxBytes`. */
  readonly readAll: (entry: ZipEntry, maxBytes?: number) => Effect.Effect<Uint8Array, SpreadsheetError>
  /** Bytes inflated (or copied, for stored entries) so far, over all reads: a work counter. */
  readonly inflatedBytes: () => number
}

const malformed = (reason: string, part?: string, offset?: number) =>
  new MalformedFile({
    format: "zip",
    reason,
    ...(part === undefined ? {} : { part }),
    ...(offset === undefined ? {} : { offset })
  })

const normalize = (name: string): string => (name.startsWith("/") ? name.slice(1) : name).toLowerCase()

interface Located {
  readonly entry: ZipEntry
  readonly dataStart: number
  readonly dataEnd: number
}

/** Checks the local headers and the layout; returns where each entry's data is. */
const locate = (bytes: Uint8Array, entries: Directory): Map<ZipEntry, Located> | MalformedFile => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const names = new TextDecoder("utf-8")
  const located = new Map<ZipEntry, Located>()
  // progress: one entry per iteration (entries are bounded by maxZipEntries)
  for (const entry of entries) {
    const at = entry.localOffset
    if (at + 30 > bytes.length) return malformed("local header out of bounds", entry.name, at)
    if (view.getUint32(at, true) !== LOCAL) return malformed("bad local header signature", entry.name, at)
    if (view.getUint16(at + 8, true) !== entry.method) {
      return malformed("local header compression method differs from the directory", entry.name, at)
    }
    const nameLength = view.getUint16(at + 26, true)
    const extraLength = view.getUint16(at + 28, true)
    const dataStart = at + 30 + nameLength + extraLength
    const dataEnd = dataStart + entry.compressedSize
    if (dataEnd > bytes.length) return malformed("entry data out of bounds", entry.name, at)
    // entry data comes before the central directory (and so the EOCD)
    if (dataEnd > entries.directoryStart) return malformed("entry data overlaps the central directory", entry.name, at)
    // The local name must be the directory's name (decoded the same way); a
    // mismatch means the directory points at the wrong data.
    if (names.decode(bytes.subarray(at + 30, at + 30 + nameLength)) !== entry.name) {
      return malformed("local header name differs from the directory", entry.name, at)
    }
    located.set(entry, { entry, dataStart, dataEnd })
  }
  // No two entries may share bytes: sort by header offset and check the gaps.
  const ordered = [...located.values()].sort((a, b) => a.entry.localOffset - b.entry.localOffset)
  // progress: i increases up to ordered.length
  for (let i = 1; i < ordered.length; i++) {
    const before = ordered[i - 1]!
    const after = ordered[i]!
    if (before.dataEnd > after.entry.localOffset) {
      return malformed("entries overlap", after.entry.name, after.entry.localOffset)
    }
  }
  return located
}

/** Opens an archive. `inflate` is replaceable for other inflate implementations (and tests). */
export const open = (
  bytes: Uint8Array,
  limits: Limits,
  inflate: InflateRaw = inflateRawWeb
): Archive | SpreadsheetError => {
  const entries = readDirectory(bytes, limits)
  if (!Array.isArray(entries)) return entries as MalformedFile | LimitExceeded
  const byName = new Map<string, ZipEntry>()
  // progress: one entry per iteration
  for (const entry of entries as ReadonlyArray<ZipEntry>) {
    const key = normalize(entry.name)
    if (byName.has(key)) return malformed("duplicate entry name", entry.name, entry.localOffset)
    byName.set(key, entry)
  }
  const located = locate(bytes, entries as Directory)
  if (located instanceof MalformedFile) return located
  let inflated = 0

  /** Counts `n` more output bytes of `entry` (`soFar` after them); an error if a limit is crossed. */
  const account = (entry: ZipEntry, soFar: number, n: number): SpreadsheetError | undefined => {
    inflated += n
    if (soFar > entry.uncompressedSize) {
      return malformed(`entry inflates past its declared size of ${entry.uncompressedSize} bytes`, entry.name)
    }
    if (inflated > limits.maxDecompressedBytes) {
      return new LimitExceeded({ limit: "maxDecompressedBytes", max: limits.maxDecompressedBytes, actual: inflated })
    }
    if (soFar > RATIO_GRACE_BYTES) {
      const ratio = soFar / Math.max(1, entry.compressedSize)
      if (ratio > limits.maxCompressionRatio) {
        return new LimitExceeded({ limit: "maxCompressionRatio", max: limits.maxCompressionRatio, actual: Math.floor(ratio) })
      }
    }
    return undefined
  }

  const finish = (entry: ZipEntry, size: number, crc: number): SpreadsheetError | undefined => {
    if (size !== entry.uncompressedSize) {
      return malformed(`entry is ${size} bytes, its directory says ${entry.uncompressedSize}`, entry.name)
    }
    if (crc !== entry.crc32) return malformed("CRC-32 mismatch", entry.name)
    return undefined
  }

  const read = (entry: ZipEntry): Stream.Stream<Uint8Array, SpreadsheetError> => {
    if ((entry.flags & 0x1) !== 0) return Stream.fail(new Encrypted({ format: "zip" }))
    if (entry.method !== STORED && entry.method !== DEFLATE) {
      return Stream.fail(
        new UnsupportedFormat({ format: "zip", reason: `compression method ${entry.method} in ${entry.name}` })
      )
    }
    const place = located.get(entry)
    if (place === undefined) return Stream.fail(malformed("entry not in this archive", entry.name))
    const data = bytes.subarray(place.dataStart, place.dataEnd)
    if (entry.method === STORED && entry.compressedSize !== entry.uncompressedSize) {
      return Stream.fail(malformed("stored entry with different compressed and uncompressed sizes", entry.name))
    }
    return Stream.fromChannel(Channel.fromTransform((_, scope: Scope.Scope) =>
      Effect.gen(function*() {
        let size = 0
        let crc = CRC32_START
        let done = false
        const inflater = entry.method === DEFLATE ? inflate(data) : undefined
        if (inflater !== undefined) yield* Scope.addFinalizer(scope, Effect.sync(() => inflater.cancel()))
        let storedAt = 0
        const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Uint8Array>, SpreadsheetError> = Effect.gen(function*() {
          // progress: each iteration emits a non-empty chunk, ends, or reads
          // one more chunk from the decompressor (which consumes input)
          while (true) {
            if (done) return yield* Cause.done()
            let chunk: Uint8Array | undefined
            if (inflater === undefined) {
              if (storedAt < data.length) {
                chunk = data.subarray(storedAt, Math.min(data.length, storedAt + STORED_SLICE))
                storedAt += chunk.length
                yield* Effect.yieldNow
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
            const error = account(entry, size, chunk.length)
            if (error !== undefined) {
              done = true
              inflater?.cancel()
              return yield* Effect.fail(error)
            }
            crc = crc32Update(crc, chunk)
            return [chunk] as const
          }
        })
        return pull
      })
    ))
  }

  const readAll = (entry: ZipEntry, maxBytes?: number) =>
    Effect.gen(function*() {
      if (maxBytes !== undefined && entry.uncompressedSize > maxBytes) {
        return yield* Effect.fail(
          new LimitExceeded({ limit: "maxDecompressedBytes", max: maxBytes, actual: entry.uncompressedSize })
        )
      }
      const chunks = yield* Stream.runCollect(read(entry))
      const out = new Uint8Array(entry.uncompressedSize)
      let at = 0
      // progress: one chunk per iteration; finish() guarantees the total fits
      for (const chunk of chunks) {
        out.set(chunk, at)
        at += chunk.length
      }
      return out
    })

  return {
    entries: entries as ReadonlyArray<ZipEntry>,
    find: (name) => byName.get(normalize(name)),
    read,
    readAll,
    inflatedBytes: () => inflated
  }
}


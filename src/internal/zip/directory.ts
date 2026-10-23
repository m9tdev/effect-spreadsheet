/**
 * ZIP central-directory reader (names, sizes, offsets; no inflation). Every
 * offset and length is checked against the buffer, and the entry count
 * against `maxZipEntries`.
 *
 * @internal
 */
import { LimitExceeded, MalformedFile } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"

export interface ZipEntry {
  readonly name: string
  /** 0 = stored, 8 = deflate. */
  readonly method: number
  readonly flags: number
  readonly compressedSize: number
  readonly uncompressedSize: number
  /** Offset of the local file header. */
  readonly localOffset: number
}

const EOCD = 0x06054b50
const ZIP64_LOCATOR = 0x07064b50
const ZIP64_EOCD = 0x06064b50
const CENTRAL = 0x02014b50
const LOCAL = 0x04034b50

const malformed = (reason: string, offset?: number) =>
  new MalformedFile(offset === undefined ? { format: "zip", reason } : { format: "zip", reason, offset })

const u64 = (view: DataView, at: number): number | undefined => {
  const lo = view.getUint32(at, true)
  const hi = view.getUint32(at + 4, true)
  // Above 2^53 can't be a real offset into a JS buffer.
  if (hi >= 0x200000) return undefined
  return hi * 0x100000000 + lo
}

/** Finds the end-of-central-directory record, scanning back over a comment of up to 64 KiB. */
const findEocd = (bytes: Uint8Array, view: DataView): number => {
  const last = bytes.length - 22
  const stop = Math.max(0, last - 0xFFFF)
  // progress: i decreases from `last` to `stop`
  for (let i = last; i >= stop; i--) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4B && view.getUint32(i, true) === EOCD) return i
  }
  return -1
}

export const readDirectory = (
  bytes: Uint8Array,
  limits: Limits
): ReadonlyArray<ZipEntry> | MalformedFile | LimitExceeded => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const eocd = findEocd(bytes, view)
  if (eocd < 0) return malformed("end of central directory not found (truncated?)")
  let count = view.getUint16(eocd + 10, true)
  let size: number | undefined = view.getUint32(eocd + 12, true)
  let start: number | undefined = view.getUint32(eocd + 16, true)
  if (count === 0xFFFF || size === 0xFFFFFFFF || start === 0xFFFFFFFF) {
    const locator = eocd - 20
    if (locator < 0 || view.getUint32(locator, true) !== ZIP64_LOCATOR) {
      return malformed("ZIP64 locator missing", Math.max(0, locator))
    }
    const at = u64(view, locator + 8)
    if (at === undefined || at + 56 > locator) return malformed("ZIP64 end of central directory out of bounds", locator)
    if (view.getUint32(at, true) !== ZIP64_EOCD) return malformed("bad ZIP64 end of central directory", at)
    const count64 = u64(view, at + 32)
    size = u64(view, at + 40)
    start = u64(view, at + 48)
    if (count64 === undefined) return malformed("bad ZIP64 entry count", at)
    count = count64
  }
  if (count > limits.maxZipEntries) {
    return new LimitExceeded({ limit: "maxZipEntries", max: limits.maxZipEntries, actual: count })
  }
  if (size === undefined || start === undefined || start + size > bytes.length) {
    return malformed("central directory out of bounds", eocd)
  }
  const entries: Array<ZipEntry> = []
  const names = new TextDecoder("utf-8")
  let p = start
  const end = start + size
  // progress: p grows by at least 46 per entry; at most `count` entries
  for (let k = 0; k < count; k++) {
    if (p + 46 > end) return malformed("central directory entry out of bounds", p)
    if (view.getUint32(p, true) !== CENTRAL) return malformed("bad central directory entry signature", p)
    const flags = view.getUint16(p + 8, true)
    const method = view.getUint16(p + 10, true)
    const compressedSize = view.getUint32(p + 20, true)
    const uncompressedSize = view.getUint32(p + 24, true)
    const nameLength = view.getUint16(p + 28, true)
    const extraLength = view.getUint16(p + 30, true)
    const commentLength = view.getUint16(p + 32, true)
    const localOffset = view.getUint32(p + 42, true)
    const next = p + 46 + nameLength + extraLength + commentLength
    if (next > end) return malformed("central directory entry out of bounds", p)
    const name = names.decode(bytes.subarray(p + 46, p + 46 + nameLength))
    // ZIP64 extra fields (sizes/offsets of 0xFFFFFFFF) are left to the full zip reader.
    entries.push({ name, method, flags, compressedSize, uncompressedSize, localOffset })
    p = next
  }
  return entries
}

/**
 * The stored (uncompressed) data of an entry, or undefined if it isn't
 * stored or is out of bounds. Used to read an ODF `mimetype`.
 */
export const storedData = (bytes: Uint8Array, entry: ZipEntry): Uint8Array | undefined => {
  if (entry.method !== 0) return undefined
  const at = entry.localOffset
  if (at + 30 > bytes.length) return undefined
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (view.getUint32(at, true) !== LOCAL) return undefined
  const dataStart = at + 30 + view.getUint16(at + 26, true) + view.getUint16(at + 28, true)
  const dataEnd = dataStart + entry.compressedSize
  if (dataEnd > bytes.length) return undefined
  return bytes.subarray(dataStart, dataEnd)
}

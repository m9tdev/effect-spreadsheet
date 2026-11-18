/**
 * What an input is, found once and handed on: format detection maps it to
 * a `Format`, and `Spreadsheet.read` and `Xls.read` read what it found, so
 * the three apply the same rules and a container is opened only once per
 * read.
 *
 * A compound file is recognised by its directory: it holds a workbook if
 * it has a `Workbook` (BIFF8) or `Book` (BIFF5) stream, which is then
 * trusted to be one (its first record is only checked to name BIFF2-4,
 * when the stream is read). A bare BIFF stream has no container to vouch
 * for it, so it must start with a well-formed BOF record (`biffBof`). Bytes
 * that start with a BIFF3 or BIFF4 BOF record type without one are a broken
 * stream: detection and both readers fail with `MalformedFile`. Bytes that
 * start with the BIFF2 BOF record type without one fail only `Xls.read`;
 * detection reads them as text or unknown data, as UTF-16LE text that
 * starts with a tab begins the same way (`09 00`).
 *
 * @internal
 */
import * as Result from "effect/Result"
import { Encrypted, type LimitExceeded, MalformedFile, UnsupportedFormat } from "../SpreadsheetError.ts"
import type { Limits } from "../SpreadsheetLimits.ts"
import type { Format } from "../Workbook.ts"
import { type Cfb, type CfbEntry, type CfbError, open as openCfb } from "./cfb/cfb.ts"
import { type Meter, noMeter } from "./meter.ts"
import { biffBof, CFB_MAGIC, detectZip, plainFormat, startsWith, ZIP_MAGICS } from "./sniff.ts"

/**
 * A BIFF5/8 workbook stream. One in a compound file is read from it only
 * when the workbook is (detection doesn't read it).
 */
export interface Biff {
  readonly _tag: "Biff"
  readonly source:
    | { readonly _tag: "Bare"; readonly stream: Uint8Array }
    | { readonly _tag: "Compound"; readonly cfb: Cfb; readonly entry: CfbEntry }
}

/** A compound file holding an encrypted OOXML package (`EncryptionInfo`/`EncryptedPackage`). */
export interface EncryptedPackage {
  readonly _tag: "Encrypted"
  readonly format: "xlsx"
}

/**
 * A workbook container this library doesn't read: a compound file without
 * a workbook stream, or a bare BIFF2-4 stream. `format` is what detection
 * reports; `found` and `reason` are what reading fails with.
 */
export interface Unreadable {
  readonly _tag: "Unreadable"
  readonly format: "xls" | "cfb"
  readonly found: string
  readonly reason: string
}

/** What a compound file or a bare BIFF stream holds. */
export type XlsProbe = Biff | EncryptedPackage | Unreadable

/**
 * Anything else, by the format its content shows: a zip (by its entry
 * names) or plain bytes (by magic numbers and text). Read by format.
 */
export interface Detected {
  readonly _tag: "Detected"
  readonly format: Exclude<Format, "xls" | "cfb">
}

export type Probe = XlsProbe | Detected

const OLD_BIFF_REASON = (version: string) => `${version.toUpperCase()} (Excel 2.x-4.x) workbooks aren't supported`

/**
 * What a compound file holds: its `Workbook` or `Book` stream, an encrypted
 * OOXML package, or nothing readable. A broken compound file fails.
 */
const probeCfb = (bytes: Uint8Array, limits: Limits, meter: Meter): Result.Result<XlsProbe, CfbError> => {
  const opened = openCfb(bytes, limits, meter)
  if (Result.isFailure(opened)) return Result.fail(opened.failure)
  const cfb = opened.success
  const entry = cfb.find("Workbook") ?? cfb.find("Book")
  if (entry !== undefined && entry.type === "stream") {
    return Result.succeed({ _tag: "Biff", source: { _tag: "Compound", cfb, entry } })
  }
  if (cfb.find("EncryptedPackage") !== undefined) return Result.succeed({ _tag: "Encrypted", format: "xlsx" })
  return Result.succeed({
    _tag: "Unreadable",
    format: "cfb",
    found: "cfb",
    reason: "compound file without a Workbook or Book stream"
  })
}

/** BOF record types of BIFF2-4, by version. */
const OLD_BOFS: ReadonlyMap<number, string> = new Map([[0x0009, "biff2"], [0x0209, "biff3"], [0x0409, "biff4"]])

/** The BIFF2-4 version a stream's first record type names, if any. */
const oldBiff = (stream: Uint8Array): string | undefined =>
  stream.length >= 2 ? OLD_BOFS.get(stream[0]! | (stream[1]! << 8)) : undefined

/**
 * A bare BIFF stream, by its BOF record, or `undefined` if it doesn't start
 * with a BOF record type. A BIFF2-4 BOF record type without a well-formed
 * BOF record fails, but a BIFF2 one only if `strictBiff2`, and is
 * `undefined` otherwise.
 */
const probeBare = (bytes: Uint8Array, strictBiff2: boolean): Result.Result<XlsProbe, MalformedFile> | undefined => {
  const version = biffBof(bytes)
  if (version === "biff5/8") return Result.succeed({ _tag: "Biff", source: { _tag: "Bare", stream: bytes } })
  if (version !== undefined) {
    return Result.succeed({ _tag: "Unreadable", format: "xls", found: version, reason: OLD_BIFF_REASON(version) })
  }
  const old = oldBiff(bytes)
  return old === undefined || (old === "biff2" && !strictBiff2)
    ? undefined
    : Result.fail(new MalformedFile({ format: "xls", reason: `malformed ${old.toUpperCase()} BOF record` }))
}

/**
 * What `bytes` are. Only a container (a compound file or a zip) is opened,
 * and a broken one fails rather than giving a format. The sectors walked
 * opening a compound file, and later reading its workbook stream, go to
 * `meter`.
 */
export const probe = (
  bytes: Uint8Array,
  limits: Limits,
  meter: Meter = noMeter
): Result.Result<Probe, MalformedFile | LimitExceeded> => {
  if (bytes.length === 0) return Result.succeed({ _tag: "Detected", format: "csv" })
  if (startsWith(bytes, CFB_MAGIC)) return probeCfb(bytes, limits, meter)
  if (ZIP_MAGICS.some((magic) => startsWith(bytes, magic))) {
    const zip = detectZip(bytes, limits)
    return Result.isFailure(zip) ? Result.fail(zip.failure) : Result.succeed({ _tag: "Detected", format: zip.success })
  }
  // a malformed BIFF2 BOF may be UTF-16LE text starting with a tab: left to plain detection
  return probeBare(bytes, false) ?? Result.succeed({ _tag: "Detected", format: plainFormat(bytes) })
}

/** The format detection reports for a probe. */
export const formatOf = (probe: Probe): Format => {
  switch (probe._tag) {
    case "Biff":
      return "xls"
    case "Encrypted":
      return "cfb"
    case "Unreadable":
    case "Detected":
      return probe.format
  }
}

/** The format of `bytes`: `probe`, mapped by `formatOf`. */
export const detect = (bytes: Uint8Array, limits: Limits): Result.Result<Format, MalformedFile | LimitExceeded> =>
  Result.map(probe(bytes, limits), formatOf)

/**
 * What the XLS reader reads: the same probe, for a compound file or a bare
 * BIFF stream; anything else is no XLS file. Unlike detection, it rejects a
 * malformed BIFF2 BOF record too: the caller asked for an XLS file.
 */
export const probeXls = (
  bytes: Uint8Array,
  limits: Limits,
  meter: Meter
): Result.Result<XlsProbe, MalformedFile | LimitExceeded> => {
  if (startsWith(bytes, CFB_MAGIC)) return probeCfb(bytes, limits, meter)
  return probeBare(bytes, true) ??
    Result.fail(new MalformedFile({ format: "xls", reason: "neither a compound file nor a BIFF stream" }))
}

/**
 * The workbook stream of a `Biff` probe. A stream from a compound file is
 * read now, and fails with `UnsupportedFormat` if its first record is a
 * BIFF2-4 BOF.
 */
export const workbookStream = (biff: Biff): Result.Result<Uint8Array, CfbError | UnsupportedFormat> => {
  const source = biff.source
  if (source._tag === "Bare") return Result.succeed(source.stream)
  const stream = source.cfb.read(source.entry)
  if (Result.isFailure(stream)) return stream
  const old = oldBiff(stream.success)
  return old === undefined
    ? stream
    : Result.fail(new UnsupportedFormat({ format: old, reason: OLD_BIFF_REASON(old) }))
}

/** What reading an `Encrypted` or `Unreadable` probe fails with. */
export const unreadableError = (probe: EncryptedPackage | Unreadable): Encrypted | UnsupportedFormat =>
  probe._tag === "Encrypted"
    ? new Encrypted({ format: probe.format })
    : new UnsupportedFormat({ format: probe.found, reason: probe.reason })

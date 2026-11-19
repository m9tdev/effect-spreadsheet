/**
 * What an input is, found once and handed on: format detection maps it to
 * a `Format`, and `Spreadsheet.read` and `Xls.read` read what it found, so
 * the three apply the same rules and a container is opened only once per
 * read.
 *
 * A zip is first known by its central directory. One that looks like an
 * OOXML package (it has `_rels/.rels`, `xl/workbook.xml` or
 * `xl/workbook.bin`) is opened, and a broken one fails. Its workbook part
 * is the target of the package's `officeDocument` relationship, as the
 * readers find it: `xl/workbook.xml` is XLSX, `xl/workbook.bin` XLSB, and
 * a part elsewhere goes by its content type. Without such a relationship
 * (or with a target the package lacks), `xl/workbook.bin` makes it XLSB and
 * `xl/workbook.xml` XLSX. A zip with neither is ODS if its `mimetype` entry
 * says so, and else just a zip.
 *
 * A compound file is recognised by its directory: it holds a workbook if
 * it has a `Workbook` (BIFF8) or `Book` (BIFF5) stream, which is then
 * trusted to be one (its first record is only checked to name BIFF2-4,
 * when the stream is read). A bare BIFF stream has no container to vouch
 * for it, so it must start with a well-formed BOF record (`biffBof`).
 * Bytes that start with a BIFF2-4 BOF record type without one fail only
 * `Xls.read`, with `MalformedFile`: detection reads them as text or unknown
 * data, as UTF-16LE text without a BOM can begin the same way (a tab is
 * `09 00`, `Љ` is `09 04`).
 *
 * @internal
 */
import { Effect } from "effect"
import * as Result from "effect/Result"
import {
  Encrypted,
  type LimitExceeded,
  MalformedFile,
  type SpreadsheetError,
  UnsupportedFormat
} from "../SpreadsheetError.ts"
import type { Limits } from "../SpreadsheetLimits.ts"
import type { Format, UnreadableFormat } from "../Workbook.ts"
import { type Cfb, type CfbEntry, type CfbError, open as openCfb } from "./cfb/cfb.ts"
import { type Meter, noMeter } from "./meter.ts"
import { contentType, lookup, type Named, packageRels } from "./opc.ts"
import { fixedWorkbook } from "./package.ts"
import { biffBof, type BiffVersion, CFB_MAGIC, plainFormat, startsWith, ZIP_MAGICS } from "./sniff.ts"
import { type Archive, locateEntry, openDirectory } from "./zip/archive.ts"
import { type Directory, normalizeName, readDirectory, type ZipEntry } from "./zip/directory.ts"

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
  readonly found: "cfb" | OldBiff
  readonly reason: string
}

/** What a compound file or a bare BIFF stream holds. */
export type XlsProbe = Biff | EncryptedPackage | Unreadable

/**
 * An OOXML package, opened: its kind and its workbook part (present, or a
 * part the package lacks, which reading reports).
 */
export interface Package {
  readonly _tag: "Package"
  readonly kind: "xlsx" | "xlsb"
  readonly archive: Archive
  readonly workbook: Named
}

/**
 * Anything else, by the format its content shows: a zip that is no
 * package (an ODF `mimetype`, or nothing known) or plain bytes (by magic
 * numbers and text). Read by format.
 */
export interface Detected {
  readonly _tag: "Detected"
  readonly format: "csv" | Exclude<UnreadableFormat, "cfb">
}

export type Probe = XlsProbe | Package | Detected

/** The BIFF versions of Excel 2.x-4.x, which aren't read. */
type OldBiff = Exclude<BiffVersion, "biff5/8">

const OLD_BIFF_REASON = (version: OldBiff) => `${version.toUpperCase()} (Excel 2.x-4.x) workbooks aren't supported`

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
const OLD_BOFS: ReadonlyMap<number, OldBiff> = new Map([[0x0009, "biff2"], [0x0209, "biff3"], [0x0409, "biff4"]])

/** The BIFF2-4 version a stream's first record type names, if any. */
const oldBiff = (stream: Uint8Array): OldBiff | undefined =>
  stream.length >= 2 ? OLD_BOFS.get(stream[0]! | (stream[1]! << 8)) : undefined

/**
 * A bare BIFF stream, by its BOF record, or `undefined` if it doesn't start
 * with a well-formed one. If `strict`, a BIFF2-4 BOF record type without a
 * well-formed BOF record fails instead.
 */
const probeBare = (bytes: Uint8Array, strict: boolean): Result.Result<XlsProbe, MalformedFile> | undefined => {
  const version = biffBof(bytes)
  if (version === "biff5/8") return Result.succeed({ _tag: "Biff", source: { _tag: "Bare", stream: bytes } })
  if (version !== undefined) {
    return Result.succeed({ _tag: "Unreadable", format: "xls", found: version, reason: OLD_BIFF_REASON(version) })
  }
  if (!strict) return undefined
  const old = oldBiff(bytes)
  return old === undefined
    ? undefined
    : Result.fail(new MalformedFile({ format: "xls", reason: `malformed ${old.toUpperCase()} BOF record` }))
}

const ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet"

/** The content types of a workbook part, by the package kind they make. */
const WORKBOOK_TYPES: ReadonlyMap<string, "xlsx" | "xlsb"> = new Map([
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "xlsx"],
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml", "xlsx"],
  ["application/vnd.ms-excel.sheet.macroenabled.main+xml", "xlsx"],
  ["application/vnd.ms-excel.template.macroenabled.main+xml", "xlsx"],
  ["application/vnd.ms-excel.addin.macroenabled.main+xml", "xlsx"],
  ["application/vnd.ms-excel.sheet.binary.macroenabled.main", "xlsb"]
])

/** The package kind of a workbook part found through the package relationships, if it is a workbook. */
const partKind = (
  archive: Archive,
  target: Extract<Named, { readonly _tag: "Part" }>,
  limits: Limits
): Effect.Effect<"xlsx" | "xlsb" | undefined, SpreadsheetError> => {
  const name = normalizeName(target.entry.name)
  if (name === fixedWorkbook.xlsx) return Effect.succeed("xlsx")
  if (name === fixedWorkbook.xlsb) return Effect.succeed("xlsb")
  return Effect.map(
    contentType(archive, target.entry, limits),
    (type) => type === undefined ? undefined : WORKBOOK_TYPES.get(type.split(";", 1)[0]!.trim().toLowerCase())
  )
}

/**
 * The workbook of an opened package-looking zip (see the module doc), or
 * `undefined` if it has none.
 */
const probePackage = (archive: Archive, limits: Limits): Effect.Effect<Package | undefined, SpreadsheetError> =>
  Effect.flatMap(packageRels(archive, limits), (rels) => {
    const target = rels.named("officeDocument")
    if (target?._tag === "Part") {
      return Effect.map(
        partKind(archive, target, limits),
        (kind) => kind === undefined ? undefined : { _tag: "Package", kind, archive, workbook: target }
      )
    }
    const kind = archive.find(fixedWorkbook.xlsb) !== undefined
      ? "xlsb"
      : archive.find(fixedWorkbook.xlsx) !== undefined
      ? "xlsx"
      : undefined
    return Effect.succeed(
      kind === undefined ? undefined : { _tag: "Package", kind, archive, workbook: target ?? lookup(archive, fixedWorkbook[kind]) }
    )
  })

/**
 * Whether one of a zip's `mimetype` entries names an ODF spreadsheet (or
 * template); the last one that reads decides. An entry reads if it is
 * stored, with the local header the archive would accept, and short.
 */
const isOds = (bytes: Uint8Array, directory: Directory, mimetypes: ReadonlyArray<ZipEntry>): boolean => {
  let mimetype: string | undefined
  // progress: one entry per iteration, entries bounded by maxZipEntries
  for (const entry of mimetypes) {
    if (entry.method !== 0) continue // not stored: no need to look at its local header
    const located = locateEntry(bytes, directory, entry)
    if (Result.isFailure(located) || located.success.kind !== "stored") continue
    const { start, end } = located.success.data
    if (end - start < 256) mimetype = new TextDecoder().decode(bytes.subarray(start, end)).trim()
  }
  return mimetype === ODS_MIME || mimetype === `${ODS_MIME}-template`
}

/**
 * What a zip is: a package (opened), an ODF spreadsheet or just a zip. A
 * broken directory, or a broken package, fails.
 */
const probeZip = (bytes: Uint8Array, limits: Limits): Effect.Effect<Package | Detected, SpreadsheetError> => {
  const read = readDirectory(bytes, limits)
  if (Result.isFailure(read)) return Effect.fail(read.failure)
  const directory = read.success
  // whether a name makes it look like an OOXML package, and the `mimetype` entries
  let looksLikePackage = false
  const mimetypes: Array<ZipEntry> = []
  // progress: one entry per iteration, entries bounded by maxZipEntries
  for (const entry of directory.entries) {
    const name = normalizeName(entry.name)
    if (name === "_rels/.rels" || name === fixedWorkbook.xlsx || name === fixedWorkbook.xlsb) looksLikePackage = true
    else if (name === "mimetype") mimetypes.push(entry)
  }
  const other = (): Detected => ({
    _tag: "Detected",
    format: mimetypes.length > 0 && isOds(bytes, directory, mimetypes) ? "ods" : "zip"
  })
  if (!looksLikePackage) return Effect.succeed(other())
  const opened = openDirectory(bytes, directory, limits)
  if (Result.isFailure(opened)) return Effect.fail(opened.failure)
  return Effect.map(probePackage(opened.success, limits), (found) => found ?? other())
}

/**
 * What `bytes` are. Only a container (a compound file or a zip) is opened,
 * and a broken one fails rather than giving a format; a zip that looks
 * like a package has its package relationships read. The sectors walked
 * opening a compound file, and later reading its workbook stream, go to
 * `meter`; the bytes inflated from a zip go to the `WorkMeter` in context.
 */
export const probe = (
  bytes: Uint8Array,
  limits: Limits,
  meter: Meter = noMeter
): Effect.Effect<Probe, SpreadsheetError> => {
  if (bytes.length === 0) return Effect.succeed({ _tag: "Detected", format: "csv" })
  if (startsWith(bytes, CFB_MAGIC)) return Effect.fromResult(probeCfb(bytes, limits, meter))
  if (ZIP_MAGICS.some((magic) => startsWith(bytes, magic))) return probeZip(bytes, limits)
  // a malformed BIFF2-4 BOF may be UTF-16LE text: left to plain detection
  const bare = probeBare(bytes, false)
  return bare === undefined ? Effect.succeed({ _tag: "Detected", format: plainFormat(bytes) }) : Effect.fromResult(bare)
}

/** The format detection reports for a probe. */
export const formatOf = (probe: Probe): Format => {
  switch (probe._tag) {
    case "Biff":
      return "xls"
    case "Encrypted":
      return "cfb"
    case "Package":
      return probe.kind
    case "Unreadable":
    case "Detected":
      return probe.format
  }
}

/** The format of `bytes`: `probe`, mapped by `formatOf`. */
export const detect = (bytes: Uint8Array, limits: Limits): Effect.Effect<Format, SpreadsheetError> =>
  Effect.map(probe(bytes, limits), formatOf)

/**
 * What the XLS reader reads: the same probe, for a compound file or a bare
 * BIFF stream; anything else is no XLS file. Unlike detection, it rejects
 * bytes that start with a BIFF2-4 BOF record type but no well-formed BOF
 * record: the caller asked for an XLS file.
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

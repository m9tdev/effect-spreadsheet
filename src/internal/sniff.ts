/**
 * Format detection from content only (magic bytes and structure), for the
 * input probe: magic numbers, a zip's entry names, a bare BIFF BOF record,
 * and what text looks like. File names and extensions are never looked at.
 *
 * @internal
 */
import * as Result from "effect/Result"
import type { LimitExceeded, MalformedFile } from "../SpreadsheetError.ts"
import type { Limits } from "../SpreadsheetLimits.ts"
import type { Format } from "../Workbook.ts"
import { detectEncoding } from "./csv/decode.ts"
import { normalizeName, readDirectory, storedData, type ZipEntry } from "./zip/directory.ts"

const TEXT_PREFIX_BYTES = 8 * 1024

export const startsWith = (bytes: Uint8Array, magic: ReadonlyArray<number>): boolean => {
  if (bytes.length < magic.length) return false
  // progress: i increases up to magic.length
  for (let i = 0; i < magic.length; i++) if (bytes[i] !== magic[i]) return false
  return true
}

/**
 * How a zip file can start: a local file header (PK 03 04), the
 * end-of-central-directory record of an empty archive (PK 05 06), the
 * marker of a spanned/split archive (PK 07 08), or the single-segment
 * marker "PK00" (PK 30 30) followed by a local file header.
 */
export const ZIP_MAGICS: ReadonlyArray<ReadonlyArray<number>> = [
  [0x50, 0x4B, 0x03, 0x04],
  [0x50, 0x4B, 0x05, 0x06],
  [0x50, 0x4B, 0x07, 0x08],
  [0x50, 0x4B, 0x30, 0x30, 0x50, 0x4B, 0x03, 0x04]
]

export const CFB_MAGIC: ReadonlyArray<number> = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
const ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet"

/** What a zip can be. */
export type ZipFormat = "xlsx" | "xlsb" | "ods" | "zip"

/** What bytes that are neither a container nor a bare BIFF stream can be. */
export type PlainFormat = Exclude<Format, ZipFormat | "xls" | "cfb">

/** The format of a zip, from its central directory; a broken directory fails. */
export const detectZip = (bytes: Uint8Array, limits: Limits): Result.Result<ZipFormat, MalformedFile | LimitExceeded> => {
  const directory = readDirectory(bytes, limits)
  return Result.isFailure(directory) ? Result.fail(directory.failure) : Result.succeed(zipFormat(bytes, directory.success.entries))
}

/** What a zip archive holds, from its entry names (and an ODF `mimetype`). */
const zipFormat = (bytes: Uint8Array, entries: ReadonlyArray<ZipEntry>): ZipFormat => {
  let xlsxPart = false
  let xlsbPart = false
  let mimetype: string | undefined
  // progress: one entry per iteration, entries bounded by maxZipEntries
  for (const entry of entries) {
    const name = normalizeName(entry.name)
    if (name === "xl/workbook.xml") xlsxPart = true
    else if (name === "xl/workbook.bin") xlsbPart = true
    else if (name === "mimetype") {
      const data = storedData(bytes, entry)
      if (data !== undefined && data.length < 256) mimetype = new TextDecoder().decode(data).trim()
    }
  }
  if (xlsbPart) return "xlsb"
  if (xlsxPart) return "xlsx"
  if (mimetype === ODS_MIME || mimetype === `${ODS_MIME}-template`) return "ods"
  return "zip"
}

/** BOF "type of the following data": workbook globals, VB module, worksheet, chart, macro sheet, BIFF4 workbook. */
const BOF_DATA_TYPES = new Set([0x0005, 0x0006, 0x0010, 0x0020, 0x0040, 0x0100])

/** The BIFF version a bare stream's BOF record names. */
export type BiffVersion = "biff2" | "biff3" | "biff4" | "biff5/8"

/**
 * The version of a raw BIFF BOF record with the length its version uses
 * and a known data type: BIFF2 (0x0009, length 4), BIFF3/4 (0x0209/0x0409,
 * length 6), BIFF5/8 (0x0809, length 8 or 16, version word 0x05xx/0x0600);
 * `undefined` for anything else. The strict check keeps UTF-16LE text
 * starting with tabs (`09 00 09 00`) out.
 */
export const biffBof = (bytes: Uint8Array): BiffVersion | undefined => {
  if (bytes.length < 8) return undefined
  const word = (at: number) => bytes[at]! | (bytes[at + 1]! << 8)
  const type = word(0)
  const length = word(2)
  if (!BOF_DATA_TYPES.has(word(6))) return undefined
  switch (type) {
    case 0x0009:
      return length === 4 ? "biff2" : undefined
    case 0x0209:
      return length === 6 ? "biff3" : undefined
    case 0x0409:
      return length === 6 ? "biff4" : undefined
    case 0x0809: {
      const version = word(4)
      return (length === 8 || length === 16) && ((version & 0xFF00) === 0x0500 || version === 0x0600) ? "biff5/8" : undefined
    }
    default:
      return undefined
  }
}

/**
 * Starts like markup: `<?`/`<!`, or a tag whose first line has no delimiter
 * outside tags. `<b>,x` is a CSV whose first field happens to start with `<`.
 */
const looksLikeMarkup = (text: string): boolean => {
  if (!text.startsWith("<")) return false
  if (text.startsWith("<?") || text.startsWith("<!")) return true
  if (!/^<\/?[A-Za-z][\w:.-]*(?:\s[^<>]*)?\/?>/.test(text)) return false
  const firstLine = text.slice(0, 4096).split(/\r\n|\r|\n/, 1)[0]!
  return !/[,;\t|]/.test(firstLine.replace(/<[^<>]*>/g, ""))
}

/** The decoded start of a text file, BOM and leading whitespace removed. */
const textPrefix = (bytes: Uint8Array): string => {
  const head = bytes.subarray(0, TEXT_PREFIX_BYTES)
  const encoding = detectEncoding(head, head.length === bytes.length)
  return new TextDecoder(encoding).decode(head).replace(/^\uFEFF?\s*/, "")
}

/** SYLK starts with an `ID;P` record and has `C;`/`F;`/`B;`/`P;` records or `E` at the end. */
const isSylk = (text: string): boolean =>
  text.startsWith("ID;P") && /\r?\n(?:[CFBP];|E\s*$)/.test(text)

/** The format of bytes that are neither a compound file, a zip nor a bare BIFF stream. */
export const plainFormat = (bytes: Uint8Array): PlainFormat => {
  if (startsWith(bytes, [0x25, 0x50, 0x44, 0x46])) return "pdf"
  if (
    startsWith(bytes, [0x89, 0x50, 0x4E, 0x47]) || startsWith(bytes, [0xFF, 0xD8, 0xFF]) ||
    startsWith(bytes, [0x47, 0x49, 0x46, 0x38]) || startsWith(bytes, [0x49, 0x49, 0x2A, 0x00]) ||
    startsWith(bytes, [0x4D, 0x4D, 0x00, 0x2A])
  ) return "image"
  const text = textPrefix(bytes)
  if (text.startsWith("{\\rtf")) return "rtf"
  if (looksLikeMarkup(text)) {
    if (text.includes("urn:schemas-microsoft-com:office:spreadsheet") || /progid\s*=\s*"Excel\.Sheet"/i.test(text)) {
      return "spreadsheetml"
    }
    if (/^<(?:!doctype\s+html|html|table|head|body|meta)\b/i.test(text) || /<html\b/i.test(text.slice(0, 1024))) {
      return "html"
    }
    return "xml"
  }
  if (isSylk(text)) return "sylk"
  // NULs in decoded text: binary data, not delimited text.
  if (text.includes("\u0000")) return "unknown"
  return "csv"
}

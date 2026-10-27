/**
 * Legacy XLS reader: BIFF8 (Excel 97-2003) and BIFF5 (Excel 5/95)
 * workbooks, in a compound file (or, rarely, as a bare workbook stream).
 *
 * ```ts
 * import * as Xls from "effect-spreadsheet/xls"
 * import { Workbook } from "effect-spreadsheet"
 *
 * const workbook = yield* Xls.read(bytes)
 * workbook.sheets        // names and visibility, in workbook order
 * workbook.definedNames  // metadata only (BIFF8)
 * Workbook.rows(workbook, 0)          // Stream<Row, SpreadsheetError>
 * Workbook.sheetDetails(workbook, 0)  // hidden columns, merges, dimension
 * ```
 *
 * Reading opens the container and reads the workbook globals: sheets,
 * code page, date system, number formats and cell styles, shared strings,
 * defined names. Each `Workbook.rows` run walks that sheet's records in
 * bounded batches. Hidden rows are emitted with `hidden: true`.
 *
 * Values: text (shared strings, labels, a formula's string result),
 * numbers (NUMBER, RK, MULRK, formula results), booleans and error codes;
 * a number whose cell style has a date format is a `Date` (1900 or 1904
 * system), as in XLSX. `numFmt` is set from the cell style. Formula text is
 * not decompiled, so cells have no `formula`. Defined names give their
 * formula for single 3-D references and areas, as Excel writes them in
 * XLSX (`Sheet!$A$1:$B$2`, `'My Sheet'!$1:$2`, `[1]Data!$A:$A` for an
 * external workbook, `#REF!$A$1` for a deleted sheet), `""` otherwise;
 * BIFF5 workbooks give no defined names.
 *
 * BIFF2, BIFF3 and BIFF4 files (Excel 2.x-4.x), bare or in a compound
 * file, fail with `UnsupportedFormat` naming the version; a password-protected workbook
 * (FILEPASS) or an encrypted OOXML package in a compound file fails with
 * `Encrypted`.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import { open as openCfb } from "../internal/cfb/cfb.ts"
import { readGlobals } from "../internal/xls/globals.ts"
import { sheetDetails, sheetRows } from "../internal/xls/sheet.ts"
import { Encrypted, LimitExceeded, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../SpreadsheetError.ts"
import * as SpreadsheetLimits from "../SpreadsheetLimits.ts"
import * as Workbook from "../Workbook.ts"

/**
 * @since 0.1.0
 * @category models
 */
export interface ReadOptions {
  /** Overrides on top of the `SpreadsheetLimits` in context. */
  readonly limits?: SpreadsheetLimits.LimitOverrides | undefined
}

const CFB_MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]

/** BOF record types of BIFF2-4, by version. */
const OLD_BOFS: ReadonlyMap<number, string> = new Map([[0x0009, "biff2"], [0x0209, "biff3"], [0x0409, "biff4"]])

/** `UnsupportedFormat` if a stream starts with a BIFF2, BIFF3 or BIFF4 BOF. */
const oldBiff = (stream: Uint8Array): UnsupportedFormat | undefined => {
  const old = stream.length >= 2 ? OLD_BOFS.get(stream[0]! | (stream[1]! << 8)) : undefined
  return old === undefined
    ? undefined
    : new UnsupportedFormat({ format: old, reason: `${old.toUpperCase()} (Excel 2.x-4.x) workbooks aren't supported` })
}

/** The workbook stream: from a compound file, or the bytes themselves when they are a bare BIFF stream. */
const workbookStream = (bytes: Uint8Array, limits: SpreadsheetLimits.Limits): Uint8Array | SpreadsheetError => {
  if (CFB_MAGIC.every((b, k) => bytes[k] === b)) {
    const cfb = openCfb(bytes, limits)
    if ("_tag" in cfb) return cfb
    const entry = cfb.find("Workbook") ?? cfb.find("Book")
    if (entry === undefined || entry.type !== "stream") {
      if (cfb.find("EncryptedPackage") !== undefined) return new Encrypted({ format: "xlsx" })
      return new UnsupportedFormat({ format: "cfb", reason: "compound file without a Workbook or Book stream" })
    }
    const stream = cfb.read(entry)
    return "_tag" in stream ? stream : oldBiff(stream) ?? stream
  }
  const old = oldBiff(bytes)
  if (old !== undefined) return old
  if (bytes.length >= 2 && (bytes[0]! | (bytes[1]! << 8)) === 0x0809) return bytes // a bare BIFF5/8 workbook stream
  return new MalformedFile({ format: "xls", reason: "neither a compound file nor a BIFF stream" })
}

/**
 * Opens an XLS workbook from its bytes.
 *
 * @since 0.1.0
 * @category readers
 */
export const read = (bytes: Uint8Array, options?: ReadOptions): Effect.Effect<Workbook.Workbook, SpreadsheetError> =>
  Effect.gen(function*() {
    const limits = yield* SpreadsheetLimits.resolve(options?.limits)
    if (bytes.length > limits.maxInputBytes) {
      return yield* new LimitExceeded({ limit: "maxInputBytes", max: limits.maxInputBytes, actual: bytes.length })
    }
    const stream = workbookStream(bytes, limits)
    if ("_tag" in stream) return yield* Effect.fail(stream)
    const globals = readGlobals(stream, limits)
    if ("_tag" in globals) return yield* Effect.fail(globals)

    const sheets: Array<Workbook.SheetInfo> = globals.sheets.map((s, index) => ({ name: s.name, index, visibility: s.visibility }))
    const definedNames: Array<Workbook.DefinedName> = globals.definedNames.map((d) => ({
      name: d.name,
      formula: d.formula,
      hidden: d.hidden,
      ...(d.scope === undefined ? {} : { scope: d.scope })
    }))
    const context = (sheet: Workbook.SheetInfo, effective: SpreadsheetLimits.Limits) => {
      const entry = globals.sheets[sheet.index]!
      return entry.kind !== "worksheet"
        ? undefined
        : { stream, offset: entry.offset, part: `sheet ${JSON.stringify(sheet.name)}`, globals, limits: effective }
    }
    return Workbook.make({
      format: "xls",
      sheets,
      definedNames,
      limits,
      // chart, macro and other sheets hold no cells
      rows: (sheet, effective) => {
        const c = context(sheet, effective)
        return c === undefined ? Stream.empty : sheetRows(c)
      },
      details: (sheet, effective) => {
        const c = context(sheet, effective)
        return c === undefined ? Effect.succeed({ hiddenColumns: [], merges: [] }) : sheetDetails(c)
      }
    })
  })

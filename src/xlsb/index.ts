/**
 * XLSB reader: Excel binary workbooks (`.xlsb`, macro-enabled or not).
 *
 * ```ts
 * import * as Xlsb from "effect-spreadsheet/xlsb"
 * import { Workbook } from "effect-spreadsheet"
 *
 * const workbook = yield* Xlsb.read(bytes)
 * workbook.sheets        // names and visibility, in workbook order
 * workbook.definedNames  // metadata only
 * Workbook.rows(workbook, "Data")     // Stream<Row, SpreadsheetError>
 * Workbook.sheetDetails(workbook, 0)  // hidden columns, merges, dimension
 * ```
 *
 * An XLSB file is a zip package like XLSX, whose parts hold BIFF12 records
 * instead of XML. Reading opens the zip, follows the package relationships
 * to the workbook part, and reads the sheet list, date system, defined
 * names, shared strings and number formats. Sheets are read lazily: each
 * `Workbook.rows` run inflates that sheet's part and walks its records as
 * a stream; `Workbook.sheetDetails` reads its hidden columns, merges and
 * declared dimension in one pass. Hidden rows are emitted with
 * `hidden: true`. Macros are ignored; chart, dialog and macro sheets have
 * no rows.
 *
 * Values: text (shared, inline and rich strings, a formula's string
 * result), numbers (real, RK, formula results), booleans and error codes;
 * a number whose cell style has a date format is a `Date` (1900 or 1904
 * system), as in XLSX, and numbers carry their `numFmt`. Formulas are
 * stored as parsed tokens and aren't decompiled, so cells have no
 * `formula`. Defined names give their formula for single 3-D references and
 * areas, as Excel writes them in XLSX (`Sheet!$A$1:$B$2`, `'My Sheet'!$1:$2`,
 * `[1]Data!$A:$A` for an external workbook, `#REF!$A$1` for a deleted
 * sheet), `""` otherwise.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import { lookup, packageRels, partRels } from "../internal/opc.ts"
import { sheetDetails, sheetRows } from "../internal/xlsb/sheet.ts"
import { type CellStyles, NO_STYLES } from "../internal/numFmt.ts"
import { readSharedStrings, readStyles } from "../internal/xlsb/parts.ts"
import { definedNames, malformed, readExternalBooks, readWorkbookBin } from "../internal/xlsb/workbook.ts"
import { open } from "../internal/zip/archive.ts"
import type { ZipEntry } from "../internal/zip/directory.ts"
import { LimitExceeded, type SpreadsheetError } from "../SpreadsheetError.ts"
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

/**
 * Opens an XLSB workbook from its bytes. Fails with `MalformedFile` for a
 * broken package or record, `LimitExceeded` when a limit is hit,
 * `Encrypted` for encrypted zip entries.
 *
 * The workbook keeps using `bytes` after this returns: its zip entries are
 * views of them, not copies, inflated when a sheet's rows are read. Don't
 * change them while the workbook is in use; pass a copy if the buffer is
 * reused.
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
    const archive = open(bytes, limits)
    if ("_tag" in archive) return yield* Effect.fail(archive)

    // the workbook part, through the package relationships
    const workbookTarget = (yield* packageRels(archive, limits)).named("officeDocument") ?? lookup(archive, "xl/workbook.bin")
    if (workbookTarget._tag === "Missing") return yield* malformed(`workbook part ${workbookTarget.name} is missing`)
    const workbookEntry = workbookTarget.entry
    const workbook = yield* readWorkbookBin(archive, workbookEntry, limits)
    const rels = yield* partRels(archive, workbookEntry, limits)

    // Two sheets may not read the same worksheet part: the inflate budget
    // counts each part once, so N sheets on one part would inflate it N
    // times for free. Excel never writes this.
    const owners = new Map<ZipEntry | string, string>()
    // progress: one sheet per iteration (bounded by maxSheets)
    for (const sheet of workbook.sheets) {
      const rel = sheet.relationshipId === undefined ? undefined : rels.byId(sheet.relationshipId)
      if (rel === undefined || rel.kind !== "worksheet") continue
      const target = rel.target
      if (target._tag === "External" || target._tag === "Invalid") continue
      // a present part is its entry; missing ones compare by name, case-insensitively
      const key = target._tag === "Part" ? target.entry : target.name.toLowerCase()
      const owner = owners.get(key)
      if (owner !== undefined) {
        const part = target._tag === "Part" ? target.entry.name : target.name
        return yield* malformed(`sheets ${JSON.stringify(owner)} and ${JSON.stringify(sheet.name)} share part ${part}`, workbookEntry.name)
      }
      owners.set(key, sheet.name)
    }

    const stringsEntry = rels.part("sharedStrings")
    const stylesEntry = rels.part("styles")
    const strings = stringsEntry === undefined ? undefined : yield* readSharedStrings(archive, stringsEntry, limits)
    const styles: CellStyles = stylesEntry === undefined ? NO_STYLES : yield* readStyles(archive, stylesEntry, limits)

    // the sheet names of external workbooks that defined names point into
    // (each external link part read once, however many books share it; a missing one leaves #REF!)
    const externalSheets = yield* readExternalBooks(archive, workbook, (id) => {
      const target = rels.byId(id)?.target
      return target?._tag === "Part" ? target.entry : undefined
    }, limits)

    const sheets: Array<Workbook.SheetInfo> = workbook.sheets.map((s, index) => ({ name: s.name, index, visibility: s.visibility }))

    /** The worksheet part of a sheet; `undefined` for chart, dialog and macro sheets (no cells). */
    const sheetPart = (sheet: Workbook.SheetInfo) => {
      const entry = workbook.sheets[sheet.index]!
      const rel = entry.relationshipId === undefined ? undefined : rels.byId(entry.relationshipId)
      if (rel === undefined) return malformed(`sheet ${JSON.stringify(sheet.name)} has no part`, workbookEntry.name)
      const target = rel.target
      if (target._tag === "External") {
        return malformed(`sheet ${JSON.stringify(sheet.name)} has an external target`, workbookEntry.name)
      }
      if (target._tag === "Invalid") {
        return malformed(`sheet ${JSON.stringify(sheet.name)} has a bad target: ${target.reason}`, workbookEntry.name)
      }
      if (rel.kind !== "worksheet") return undefined
      return target._tag === "Part" ? target.entry : malformed(`sheet part ${target.name} is missing`, workbookEntry.name)
    }

    return Workbook.make({
      format: "xlsb",
      sheets,
      definedNames: definedNames(workbook, externalSheets),
      limits,
      rows: (sheet, effective) => {
        const part = sheetPart(sheet)
        if (part === undefined) return Stream.empty
        if ("_tag" in part) return Stream.fail(part)
        return sheetRows(archive.read(part, effective), { part: part.name, strings, styles, date1904: workbook.date1904, limits: effective })
      },
      details: (sheet, effective) => {
        const part = sheetPart(sheet)
        if (part === undefined) return Effect.succeed({ hiddenColumns: [], merges: [] })
        if ("_tag" in part) return Effect.fail(part)
        return sheetDetails(archive.read(part, effective), part.name, effective)
      }
    })
  })

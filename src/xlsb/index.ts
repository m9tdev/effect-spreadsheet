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
import { sheetDetails, sheetRows } from "../internal/xlsb/sheet.ts"
import { readSharedStrings, readStyles } from "../internal/xlsb/parts.ts"
import { definedNames, malformed, readExternalSheets, readWorkbookBin, referencedBooks } from "../internal/xlsb/workbook.ts"
import { NO_STYLES, readRelationships, relationshipKind, type Styles } from "../internal/xlsx/parts.ts"
import { open } from "../internal/zip/archive.ts"
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
    const root = yield* readRelationships(archive, "", limits)
    const officeDocument = root.find((r) => relationshipKind(r.type) === "officeDocument" && r.part !== undefined)
    const workbookPart = officeDocument?.part ?? "xl/workbook.bin"
    const workbookEntry = archive.find(workbookPart)
    if (workbookEntry === undefined) return yield* malformed(`workbook part ${workbookPart} is missing`)
    const workbook = yield* readWorkbookBin(archive, workbookEntry, limits)
    const relationships = yield* readRelationships(archive, workbookEntry.name, limits)
    const byId = new Map(relationships.map((r) => [r.id, r]))

    // Two sheets may not read the same worksheet part: the inflate budget
    // counts each part once, so N sheets on one part would inflate it N
    // times for free. Excel never writes this.
    const owners = new Map<string, string>()
    // progress: one sheet per iteration (bounded by maxSheets)
    for (const sheet of workbook.sheets) {
      const rel = sheet.relationshipId === undefined ? undefined : byId.get(sheet.relationshipId)
      if (rel === undefined || rel.part === undefined || relationshipKind(rel.type) !== "worksheet") continue
      const part = archive.find(rel.part)?.name ?? rel.part
      const key = part.toLowerCase()
      const owner = owners.get(key)
      if (owner !== undefined) {
        return yield* malformed(`sheets ${JSON.stringify(owner)} and ${JSON.stringify(sheet.name)} share part ${part}`, workbookEntry.name)
      }
      owners.set(key, sheet.name)
    }

    const partEntry = (kind: string) => {
      const rel = relationships.find((r) => relationshipKind(r.type) === kind && r.part !== undefined)
      return rel === undefined ? undefined : archive.find(rel.part!)
    }
    const stringsEntry = partEntry("sharedStrings")
    const stylesEntry = partEntry("styles")
    const strings = stringsEntry === undefined ? undefined : yield* readSharedStrings(archive, stringsEntry, limits)
    const styles: Styles = stylesEntry === undefined ? NO_STYLES : yield* readStyles(archive, stylesEntry, limits)

    // the sheet names of external workbooks that defined names point into
    // (each external link part read once; a missing one leaves #REF!)
    const externalSheets = new Map<number, ReadonlyArray<string>>()
    // progress: one book per iteration (at most one per BrtSupBookSrc record)
    for (const book of referencedBooks(workbook)) {
      const entry = workbook.books[book]!
      const rel = entry.kind === "external" ? byId.get(entry.relationshipId) : undefined
      const linkEntry = rel?.part === undefined ? undefined : archive.find(rel.part)
      if (linkEntry !== undefined) externalSheets.set(book, yield* readExternalSheets(archive, linkEntry, limits))
    }

    const sheets: Array<Workbook.SheetInfo> = workbook.sheets.map((s, index) => ({ name: s.name, index, visibility: s.visibility }))

    /** The worksheet part of a sheet; `undefined` for chart, dialog and macro sheets (no cells). */
    const sheetPart = (sheet: Workbook.SheetInfo) => {
      const entry = workbook.sheets[sheet.index]!
      const rel = entry.relationshipId === undefined ? undefined : byId.get(entry.relationshipId)
      if (rel === undefined || rel.part === undefined) {
        return malformed(`sheet ${JSON.stringify(sheet.name)} has no part`, workbookEntry.name)
      }
      if (relationshipKind(rel.type) !== "worksheet") return undefined
      return archive.find(rel.part) ?? malformed(`sheet part ${rel.part} is missing`, workbookEntry.name)
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
        return sheetRows(archive.read(part), { part: part.name, strings, styles, date1904: workbook.date1904, limits: effective })
      },
      details: (sheet, effective) => {
        const part = sheetPart(sheet)
        if (part === undefined) return Effect.succeed({ hiddenColumns: [], merges: [] })
        if ("_tag" in part) return Effect.fail(part)
        return sheetDetails(archive.read(part), part.name, effective)
      }
    })
  })

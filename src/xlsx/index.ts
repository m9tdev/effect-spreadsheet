/**
 * XLSX / XLSM reader.
 *
 * ```ts
 * import * as Xlsx from "effect-spreadsheet/xlsx"
 * import { Workbook } from "effect-spreadsheet"
 *
 * const workbook = yield* Xlsx.read(bytes)
 * workbook.sheets        // names and visibility, in workbook order
 * workbook.definedNames  // metadata only
 * Workbook.rows(workbook, "Data") // Stream<Row, SpreadsheetError>
 * ```
 *
 * Reading opens the zip, follows the package relationships to the workbook
 * part, and reads the sheet list, defined names, date system, shared
 * strings and number formats. Sheets are read lazily: each
 * `Workbook.rows` run inflates and parses that sheet's part as a stream;
 * `Workbook.sheetDetails` reads its hidden columns, merges and declared
 * dimension in one pass.
 * Macro-enabled workbooks (`.xlsm`) read the same way; their macros are
 * ignored. Chart sheets, dialog sheets and macro sheets have no rows.
 *
 * Cell values: shared, inline and formula strings are `Text`; booleans,
 * errors (`#N/A`, ...) and numbers have their types; a number whose number
 * format shows a date or time is a `Date` (wall clock as UTC, rounded to
 * the millisecond, in the workbook's 1900 or 1904 date system), as is an
 * ISO-dated cell (`t="d"`: without a zone its wall clock as UTC, with an
 * offset the instant it names). Number formats other than General are given as
 * `numFmt`; formula text (as stored, shared formulas only on their first
 * cell) as `formula`. Formulas are not evaluated.
 *
 * @since 0.1.0
 */
import { Effect, Stream } from "effect"
import { open } from "../internal/zip/archive.ts"
import {
  malformed,
  NO_STYLES,
  readRelationships,
  readSharedStrings,
  readStyles,
  readWorkbookPart,
  relationshipKind,
  type Styles
} from "../internal/xlsx/parts.ts"
import { sheetDetails } from "../internal/xlsx/details.ts"
import { sheetRows } from "../internal/xlsx/sheet.ts"
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
 * Opens an XLSX or XLSM workbook from its bytes. Fails with
 * `MalformedFile` for a broken package, `LimitExceeded` when a limit is
 * hit, `Encrypted` for encrypted zip entries.
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
    const workbookPart = officeDocument?.part ?? "xl/workbook.xml"
    const workbookEntry = archive.find(workbookPart)
    if (workbookEntry === undefined) return yield* malformed(`workbook part ${workbookPart} is missing`)
    const workbook = yield* readWorkbookPart(archive, workbookEntry, limits)
    const relationships = yield* readRelationships(archive, workbookEntry.name, limits)
    const byId = new Map(relationships.map((r) => [r.id, r]))

    const partEntry = (kind: string) => {
      const rel = relationships.find((r) => relationshipKind(r.type) === kind && r.part !== undefined)
      return rel === undefined ? undefined : archive.find(rel.part!)
    }
    const stringsEntry = partEntry("sharedStrings")
    const stylesEntry = partEntry("styles")
    const strings = stringsEntry === undefined ? undefined : yield* readSharedStrings(archive, stringsEntry, limits)
    const styles: Styles = stylesEntry === undefined ? NO_STYLES : yield* readStyles(archive, stylesEntry, limits)

    const sheets: Array<Workbook.SheetInfo> = workbook.sheets.map((s, index) => ({
      name: s.name,
      index,
      visibility: s.visibility
    }))
    const definedNames: Array<Workbook.DefinedName> = workbook.definedNames.map((d) => ({
      name: d.name,
      formula: d.formula,
      hidden: d.hidden,
      ...(d.scope !== undefined && d.scope < sheets.length ? { scope: d.scope } : {})
    }))

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
      format: "xlsx",
      sheets,
      definedNames,
      limits,
      rows: (sheet, effective) => {
        const part = sheetPart(sheet)
        if (part === undefined) return Stream.empty
        if ("_tag" in part) return Stream.fail(part)
        return sheetRows(archive.read(part), {
          part: part.name,
          strings,
          styles,
          date1904: workbook.date1904,
          limits: effective
        })
      },
      details: (sheet, effective) => {
        const part = sheetPart(sheet)
        if (part === undefined) return Effect.succeed({ hiddenColumns: [], merges: [] })
        if ("_tag" in part) return Effect.fail(part)
        return sheetDetails(archive.read(part), part.name, effective)
      }
    })
  })

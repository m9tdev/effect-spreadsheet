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
import { type CellStyles, NO_STYLES } from "../internal/numFmt.ts"
import { open } from "../internal/zip/archive.ts"
import type { ZipEntry } from "../internal/zip/directory.ts"
import { lookup, packageRels, partRels } from "../internal/opc.ts"
import { malformed, readSharedStrings, readStyles, readWorkbookPart } from "../internal/xlsx/parts.ts"
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
    const workbookTarget = (yield* packageRels(archive, limits)).named("officeDocument") ?? lookup(archive, "xl/workbook.xml")
    if (workbookTarget._tag === "Missing") return yield* malformed(`workbook part ${workbookTarget.name} is missing`)
    const workbookEntry = workbookTarget.entry
    const workbook = yield* readWorkbookPart(archive, workbookEntry, limits)
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
      format: "xlsx",
      sheets,
      definedNames,
      limits,
      rows: (sheet, effective) => {
        const part = sheetPart(sheet)
        if (part === undefined) return Stream.empty
        if ("_tag" in part) return Stream.fail(part)
        return sheetRows(archive.read(part, effective), {
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
        return sheetDetails(archive.read(part, effective), part.name, effective)
      }
    })
  })

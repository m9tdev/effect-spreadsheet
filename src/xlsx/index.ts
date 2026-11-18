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
import { Effect } from "effect"
import { accept } from "../internal/input.ts"
import { readXlsx } from "../internal/xlsx/read.ts"
import { type SpreadsheetError } from "../SpreadsheetError.ts"
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
  // suspend: the options are read when the effect runs, each time it runs
  Effect.suspend(() => accept(bytes, options?.limits, readXlsx))

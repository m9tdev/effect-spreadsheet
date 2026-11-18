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
import { Effect } from "effect"
import { accept } from "../internal/input.ts"
import { readXlsb } from "../internal/xlsb/read.ts"
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
  // suspend: the options are read when the effect runs, each time it runs
  Effect.suspend(() => accept(bytes, options?.limits, readXlsb))

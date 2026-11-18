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
import { Effect } from "effect"
import { accept } from "../internal/input.ts"
import { readXls } from "../internal/xls/read.ts"
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
 * Opens an XLS workbook from its bytes.
 *
 * The workbook keeps using `bytes` after this returns: a bare BIFF stream,
 * or a workbook stream in consecutive sectors of its compound file, is a
 * view of them, not a copy, and sheets are read from it when their rows
 * are. Don't change them while the workbook is in use; pass a copy if the
 * buffer is reused.
 *
 * @since 0.1.0
 * @category readers
 */
export const read = (bytes: Uint8Array, options?: ReadOptions): Effect.Effect<Workbook.Workbook, SpreadsheetError> =>
  Effect.flatMap(accept(bytes, options?.limits), readXls)

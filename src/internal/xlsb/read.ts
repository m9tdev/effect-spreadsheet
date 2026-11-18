/**
 * The XLSB reader behind `Xlsb.read` and `Spreadsheet.read`, for an input
 * already accepted (limits resolved, size checked).
 */
import { Effect } from "effect"
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Workbook } from "../../Workbook.ts"
import type { Accepted } from "../input.ts"
import { type PackageFormat, readPackage } from "../package.ts"
import { open } from "../zip/archive.ts"
import { readSharedStrings, readStyles } from "./parts.ts"
import { sheetDetails, sheetRows } from "./sheet.ts"
import { definedNames, malformed, readExternalBooks, readWorkbookBin, type WorkbookBin } from "./workbook.ts"

const XLSB: PackageFormat<WorkbookBin> = {
  format: "xlsb",
  fallbackWorkbook: "xl/workbook.bin",
  malformed,
  readWorkbook: readWorkbookBin,
  readStrings: readSharedStrings,
  readStyles,
  // the sheet names of external workbooks that defined names point into
  // (each external link part read once, however many books share it; a missing one leaves #REF!)
  definedNames: (workbook, rels, archive, limits) =>
    Effect.map(
      readExternalBooks(archive, workbook, (id) => {
        const target = rels.byId(id)?.target
        return target?._tag === "Part" ? target.entry : undefined
      }, limits),
      (externalSheets) => definedNames(workbook, externalSheets)
    ),
  rows: sheetRows,
  details: sheetDetails
}

/** Opens the zip package and reads the workbook. */
export const readXlsb = ({ bytes, limits }: Accepted): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.flatMap(Effect.fromResult(open(bytes, limits)), (archive) => readPackage(archive, limits, XLSB))

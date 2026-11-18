/**
 * The XLSX reader behind `Xlsx.read` and `Spreadsheet.read`, for an input
 * already accepted (limits resolved, size checked).
 */
import { Effect } from "effect"
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Workbook } from "../../Workbook.ts"
import { type Accepted, stored } from "../input.ts"
import { type PackageFormat, readPackage } from "../package.ts"
import { toDefinedNames } from "../workbook.ts"
import { open } from "../zip/archive.ts"
import { sheetDetails } from "./details.ts"
import { malformed, readSharedStrings, readStyles, readWorkbookPart, type WorkbookPart } from "./parts.ts"
import { sheetRows } from "./sheet.ts"

const XLSX: PackageFormat<WorkbookPart> = {
  format: "xlsx",
  fallbackWorkbook: "xl/workbook.xml",
  malformed,
  readWorkbook: readWorkbookPart,
  readStrings: readSharedStrings,
  readStyles,
  definedNames: (workbook) => Effect.succeed(toDefinedNames(workbook.definedNames, workbook.sheets.length)),
  rows: sheetRows,
  details: sheetDetails
}

/** Opens the zip package and reads the workbook. */
export const readXlsx = ({ bytes, limits }: Accepted): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.flatMap(Effect.fromResult(open(bytes, limits)), (archive) => readPackage(archive, stored(limits), XLSX))

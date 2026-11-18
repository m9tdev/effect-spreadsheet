/**
 * The XLSX reader behind `Xlsx.read` and `Spreadsheet.read`, for an input
 * already accepted (limits resolved, size checked).
 */
import { Effect } from "effect"
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Workbook } from "../../Workbook.ts"
import { type Accepted, stored, type ValidLimits } from "../input.ts"
import type { Named } from "../opc.ts"
import { fixedWorkbook, type PackageFormat, readPackage, workbookPart } from "../package.ts"
import { toDefinedNames } from "../workbook.ts"
import { type Archive, open } from "../zip/archive.ts"
import { sheetDetails } from "./details.ts"
import { malformed, readSharedStrings, readStyles, readWorkbookPart, type WorkbookPart } from "./parts.ts"
import { sheetRows } from "./sheet.ts"

const XLSX: PackageFormat<WorkbookPart> = {
  format: "xlsx",
  malformed,
  readWorkbook: readWorkbookPart,
  readStrings: readSharedStrings,
  readStyles,
  definedNames: (workbook) => Effect.succeed(toDefinedNames(workbook.definedNames, workbook.sheets.length)),
  rows: sheetRows,
  details: sheetDetails
}

/** Reads the workbook of an opened package from its workbook part. */
export const readXlsxPackage = (
  archive: Archive,
  workbook: Named,
  limits: ValidLimits
): Effect.Effect<Workbook, SpreadsheetError> => readPackage(archive, workbook, stored(limits), XLSX)

/** Opens the zip package and reads the workbook. */
export const readXlsx = ({ bytes, limits }: Accepted): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.flatMap(Effect.fromResult(open(bytes, limits)), (archive) =>
    Effect.flatMap(workbookPart(archive, fixedWorkbook.xlsx, limits), (workbook) => readXlsxPackage(archive, workbook, limits)))

/**
 * The XLS reader behind `Xls.read` and `Spreadsheet.read`, for an input
 * already accepted (limits resolved, size checked): finds the workbook
 * stream (in a compound file, or bare) and reads the workbook globals.
 */
import { Effect } from "effect"
import * as Result from "effect/Result"
import { Encrypted, MalformedFile, type SpreadsheetError, UnsupportedFormat } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Workbook } from "../../Workbook.ts"
import { open as openCfb } from "../cfb/cfb.ts"
import { type Accepted, stored } from "../input.ts"
import { type Meter, WorkMeter } from "../meter.ts"
import { fromSources } from "../workbook.ts"
import { readGlobals, type SheetEntry } from "./globals.ts"
import { type SheetContext, sheetDetails, sheetRows } from "./sheet.ts"

const CFB_MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]

/** BOF record types of BIFF2-4, by version. */
const OLD_BOFS: ReadonlyMap<number, string> = new Map([[0x0009, "biff2"], [0x0209, "biff3"], [0x0409, "biff4"]])

/** The stream, or `UnsupportedFormat` if it starts with a BIFF2, BIFF3 or BIFF4 BOF. */
const notOldBiff = (stream: Uint8Array): Result.Result<Uint8Array, UnsupportedFormat> => {
  const old = stream.length >= 2 ? OLD_BOFS.get(stream[0]! | (stream[1]! << 8)) : undefined
  return old === undefined
    ? Result.succeed(stream)
    : Result.fail(new UnsupportedFormat({ format: old, reason: `${old.toUpperCase()} (Excel 2.x-4.x) workbooks aren't supported` }))
}

/** The workbook stream: from a compound file, or the bytes themselves when they are a bare BIFF stream. */
const workbookStream = (
  bytes: Uint8Array,
  limits: Limits,
  meter: Meter
): Result.Result<Uint8Array, SpreadsheetError> => {
  if (CFB_MAGIC.every((b, k) => bytes[k] === b)) {
    const opened = openCfb(bytes, limits, meter)
    if (Result.isFailure(opened)) return Result.fail(opened.failure)
    const cfb = opened.success
    const entry = cfb.find("Workbook") ?? cfb.find("Book")
    if (entry === undefined || entry.type !== "stream") {
      if (cfb.find("EncryptedPackage") !== undefined) return Result.fail(new Encrypted({ format: "xlsx" }))
      return Result.fail(new UnsupportedFormat({ format: "cfb", reason: "compound file without a Workbook or Book stream" }))
    }
    const stream = cfb.read(entry)
    return Result.isFailure(stream) ? stream : notOldBiff(stream.success)
  }
  const old = notOldBiff(bytes)
  if (Result.isFailure(old)) return old
  if (bytes.length >= 2 && (bytes[0]! | (bytes[1]! << 8)) === 0x0809) return old // a bare BIFF5/8 workbook stream
  return Result.fail(new MalformedFile({ format: "xls", reason: "neither a compound file nor a BIFF stream" }))
}

/** Reads the workbook from its stream; sheets are read when their rows are. */
export const readXls = ({ bytes, limits }: Accepted): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.gen(function*() {
    const stream = yield* Effect.fromResult(workbookStream(bytes, limits, yield* WorkMeter))
    const globals = yield* readGlobals(stream, limits)

    const context = (entry: SheetEntry, effective: Limits): SheetContext => ({
      stream,
      offset: entry.offset,
      end: entry.end,
      part: `sheet ${JSON.stringify(entry.name)}`,
      globals,
      limits: effective
    })
    return fromSources({
      format: "xls",
      sheets: globals.sheets.map((entry) => ({
        name: entry.name,
        visibility: entry.visibility,
        // chart, macro and other sheets hold no cells
        source: entry.kind !== "worksheet" ? { _tag: "NoCells" } : {
          _tag: "Cells",
          rows: (effective) => sheetRows(context(entry, effective)),
          details: (effective) => sheetDetails(context(entry, effective))
        }
      })),
      definedNames: globals.definedNames,
      limits: stored(limits)
    })
  })

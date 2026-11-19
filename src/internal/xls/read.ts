/**
 * The XLS reader behind `Xls.read` and `Spreadsheet.read`, for an input
 * already accepted (limits resolved, size checked) and probed (a compound
 * file or a bare BIFF stream): reads the workbook stream's globals.
 */
import { Effect } from "effect"
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Workbook } from "../../Workbook.ts"
import { type Accepted, stored, type ValidLimits } from "../input.ts"
import { WorkMeter } from "../meter.ts"
import { probeXls, unreadableError, workbookStream, type XlsProbe } from "../probe.ts"
import { fromSources, type SheetSource } from "../workbook.ts"
import { readGlobals, type SheetEntry, sheetPart } from "./globals.ts"
import { type SheetContext, sheetDetails, sheetRows } from "./sheet.ts"

/**
 * Reads the workbook a probe found; sheets are read when their rows are.
 * An encrypted package or an unreadable container fails. A worksheet whose
 * BOUNDSHEET offset is no substream start fails when it is read
 * (`SheetStart`).
 */
export const readProbed = (probe: XlsProbe, limits: ValidLimits): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.gen(function*() {
    if (probe._tag !== "Biff") return yield* Effect.fail(unreadableError(probe))
    const stream = yield* Effect.fromResult(workbookStream(probe))
    const globals = yield* readGlobals(stream, limits)

    const context = (entry: SheetEntry, end: number, effective: Limits): SheetContext => ({
      stream,
      offset: entry.offset,
      end,
      part: sheetPart(entry.name),
      globals,
      limits: effective
    })
    const source = (entry: SheetEntry): SheetSource => {
      // chart, macro and other sheets hold no cells
      if (entry.kind !== "worksheet") return { _tag: "NoCells" }
      const start = entry.start
      if (start._tag === "Broken") return start
      return {
        _tag: "Cells",
        rows: (effective) => sheetRows(context(entry, start.end, effective)),
        details: (effective) => sheetDetails(context(entry, start.end, effective))
      }
    }
    return fromSources({
      format: "xls",
      sheets: globals.sheets.map((entry) => ({ name: entry.name, visibility: entry.visibility, source: source(entry) })),
      definedNames: globals.definedNames,
      limits: stored(limits)
    })
  })

/** Probes the input as a compound file or a bare BIFF stream, and reads it. */
export const readXls = ({ bytes, limits }: Accepted): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.flatMap(WorkMeter, (meter) =>
    Effect.flatMap(Effect.fromResult(probeXls(bytes, limits, meter)), (probe) => readProbed(probe, limits)))

/** Fuzz tasks run inside a worker thread (see host.ts / worker.ts). */
import { Effect, Stream } from "effect"
import * as Csv from "../../src/csv/index.ts"
import * as Spreadsheet from "../../src/Spreadsheet.ts"
import type { LimitOverrides } from "../../src/SpreadsheetLimits.ts"
import * as Workbook from "../../src/Workbook.ts"

export type FuzzTask =
  | { readonly op: "read"; readonly bytes: Uint8Array; readonly limits?: LimitOverrides | undefined }
  | { readonly op: "detect"; readonly bytes: Uint8Array }
  | {
    readonly op: "csv"
    readonly chunks: ReadonlyArray<Uint8Array>
    readonly options?: Omit<Csv.ReadOptions, "sheetName"> | undefined
  }

/** The effect a task runs: Spreadsheet.read then every sheet's rows, detectFormat, or Csv.rows. */
export const effectOf = (task: FuzzTask): Effect.Effect<unknown, unknown> => {
  switch (task.op) {
    case "read":
      return Effect.flatMap(Spreadsheet.read(task.bytes, { limits: task.limits }), (wb) =>
        Effect.forEach(wb.sheets, (sheet) => Stream.runDrain(Workbook.rows(wb, sheet.index))))
    case "detect":
      return Spreadsheet.detectFormat(task.bytes)
    case "csv":
      return Stream.runDrain(Csv.rows(Stream.fromIterable(task.chunks), task.options))
  }
}

/** Fuzz tasks run inside a worker thread (see host.ts / worker.ts). */
import { Effect, Stream } from "effect"
import * as Csv from "../../src/csv/index.ts"
import * as Spreadsheet from "../../src/Spreadsheet.ts"
import { open } from "../../src/internal/zip/archive.ts"
import { parse } from "../../src/internal/xml/parse.ts"
import type { XmlHandler } from "../../src/internal/xml/tokenizer.ts"
import * as SpreadsheetLimits from "../../src/SpreadsheetLimits.ts"
import type { LimitOverrides } from "../../src/SpreadsheetLimits.ts"

const ignore: XmlHandler = { open: () => undefined, close: () => undefined, text: () => undefined }
import * as Workbook from "../../src/Workbook.ts"

export type FuzzTask =
  | { readonly op: "read"; readonly bytes: Uint8Array; readonly limits?: LimitOverrides | undefined }
  | { readonly op: "detect"; readonly bytes: Uint8Array }
  | { readonly op: "zip"; readonly bytes: Uint8Array; readonly limits?: LimitOverrides | undefined }
  | { readonly op: "xml"; readonly chunks: ReadonlyArray<Uint8Array>; readonly limits?: LimitOverrides | undefined }
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
    case "zip":
      // open the archive, inflate every entry and tokenize the XML ones
      return Effect.flatMap(SpreadsheetLimits.resolve(task.limits), (limits) => {
        const archive = open(task.bytes, limits)
        if ("_tag" in archive) return Effect.fail(archive)
        return Effect.forEach(archive.entries, (entry) =>
          /\.(xml|rels)$/i.test(entry.name)
            ? parse(archive.read(entry), ignore, limits, entry.name)
            : Stream.runDrain(archive.read(entry)), { discard: true })
      })
    case "xml":
      return Effect.flatMap(SpreadsheetLimits.resolve(task.limits), (limits) =>
        parse(Stream.fromIterable(task.chunks), ignore, limits, "fuzz.xml"))
  }
}

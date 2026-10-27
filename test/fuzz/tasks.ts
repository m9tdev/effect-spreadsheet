/** Fuzz tasks run inside a worker thread (see host.ts / worker.ts). */
import { Effect, Stream } from "effect"
import * as Csv from "../../src/csv/index.ts"
import * as Spreadsheet from "../../src/Spreadsheet.ts"
import { RecordReader } from "../../src/internal/biff/records.ts"
import { readRichString, SegmentReader } from "../../src/internal/biff/segments.ts"
import { open as openCfb } from "../../src/internal/cfb/cfb.ts"
import { open } from "../../src/internal/zip/archive.ts"
import { parse } from "../../src/internal/xml/parse.ts"
import type { XmlHandler } from "../../src/internal/xml/tokenizer.ts"
import * as SpreadsheetLimits from "../../src/SpreadsheetLimits.ts"
import type { LimitOverrides } from "../../src/SpreadsheetLimits.ts"
import * as Workbook from "../../src/Workbook.ts"

const ignore: XmlHandler = { open: () => undefined, close: () => undefined, text: () => undefined }

export type FuzzTask =
  | { readonly op: "read"; readonly bytes: Uint8Array; readonly limits?: LimitOverrides | undefined }
  | { readonly op: "detect"; readonly bytes: Uint8Array }
  | { readonly op: "zip"; readonly bytes: Uint8Array; readonly limits?: LimitOverrides | undefined }
  | { readonly op: "xml"; readonly chunks: ReadonlyArray<Uint8Array>; readonly limits?: LimitOverrides | undefined }
  | { readonly op: "cfb"; readonly bytes: Uint8Array; readonly limits?: LimitOverrides | undefined }
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
    case "cfb":
      // open the container, read every stream, walk the workbook stream's
      // records and decode its SST strings
      return Effect.flatMap(SpreadsheetLimits.resolve(task.limits), (limits) =>
        Effect.suspend(() => {
          const cfb = openCfb(task.bytes, limits)
          if ("_tag" in cfb) return Effect.fail(cfb)
          for (const entry of cfb.entries) {
            if (entry.type !== "stream") continue
            const data = cfb.read(entry)
            if ("_tag" in data) return Effect.fail(data)
            if (!/^(workbook|book)$/i.test(entry.path)) continue
            const reader = new RecordReader(data, entry.path)
            // progress: every record advances the reader
            for (;;) {
              const record = reader.nextWithContinues(limits.maxDecompressedBytes, "maxDecompressedBytes")
              if (record === undefined) break
              if ("_tag" in record) return Effect.fail(record)
              if (record.type !== 0x00FC) continue
              const strings = new SegmentReader(record.segments)
              strings.skip(8)
              // progress: each string read consumes at least 3 bytes or fails
              while (!strings.failed && strings.remaining() > 0) {
                if (readRichString(strings, limits.maxStringLength) === undefined) break
              }
            }
          }
          return Effect.void
        }))
    case "xml":
      return Effect.flatMap(SpreadsheetLimits.resolve(task.limits), (limits) =>
        parse(Stream.fromIterable(task.chunks), ignore, limits, "fuzz.xml"))
  }
}

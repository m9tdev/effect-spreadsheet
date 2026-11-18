/**
 * What every entry point checks before reading: the limits for the call
 * (context plus overrides, validated) and the input size; and how
 * Workbook.rows / sheetDetails apply a call's overrides on top of the
 * workbook's limits.
 */
import { assert, describe, it } from "@effect/vitest"
import { Cause, Effect, Exit, Stream } from "effect"
import * as Csv from "../src/csv/index.ts"
import { type ValidLimits, mergeValid, resolveLimits } from "../src/internal/input.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import type { SpreadsheetError } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import * as Xlsb from "../src/xlsb/index.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { enc } from "./helpers/run.ts"
import * as X from "./helpers/xls.ts"
import * as B from "./helpers/xlsb.ts"
import { xlsx } from "./helpers/xlsx.ts"

type Read = (bytes: Uint8Array, options?: { readonly limits?: SpreadsheetLimits.LimitOverrides }) => Effect.Effect<Workbook.Workbook, SpreadsheetError>

/** Three rows of one number each, in every readable format. */
const inputs = Effect.promise(async () => ({
  csv: enc("1\n2\n3\n"),
  xlsx: await xlsx({ sheets: [{ name: "S", data: `<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>` }] }),
  xlsb: await B.xlsb({
    sheets: [{ name: "S", part: B.sheetBin([B.rowHdr(0), B.rk(0, B.rkInt(1)), B.rowHdr(1), B.rk(0, B.rkInt(2)), B.rowHdr(2), B.rk(0, B.rkInt(3))]) }]
  }),
  xls: X.xls({ sheets: [{ name: "S", records: [X.number(0, 0, 0, 1), X.number(1, 0, 0, 2), X.number(2, 0, 0, 3)] }] })
}))

/** Every public reader of bytes, with the format it reads. */
const readers: ReadonlyArray<readonly [name: string, format: "csv" | "xlsx" | "xlsb" | "xls", read: Read]> = [
  ["Csv.read", "csv", Csv.read],
  ["Xlsx.read", "xlsx", Xlsx.read],
  ["Xlsb.read", "xlsb", Xlsb.read],
  ["Xls.read", "xls", Xls.read],
  ["Spreadsheet.read (csv)", "csv", Spreadsheet.read],
  ["Spreadsheet.read (xlsx)", "xlsx", Spreadsheet.read],
  ["Spreadsheet.read (xlsb)", "xlsb", Spreadsheet.read],
  ["Spreadsheet.read (xls)", "xls", Spreadsheet.read]
]

/** The defect's message, or a failed assertion when `effect` doesn't die. */
const dies = <A, E, R>(effect: Effect.Effect<A, E, R>) =>
  Effect.gen(function*() {
    const exit = yield* Effect.exit(effect)
    assert.isTrue(Exit.isFailure(exit) && Cause.hasDies(exit.cause), `expected a defect, got ${exit}`)
    return Exit.isFailure(exit) ? String(Cause.squash(exit.cause)) : ""
  })

const values = (rows: ReadonlyArray<Workbook.Row>) => rows.map((row) => row.cells.map((cell) => (cell._tag === "Empty" ? "" : String(cell.value))))

describe("input: limits and size, once at the entry point", () => {
  it.effect("every reader and detectFormat check maxInputBytes before reading, against the call's limit", () =>
    Effect.gen(function*() {
      const bytes = yield* inputs
      for (const [name, format, read] of readers) {
        const input = bytes[format]
        const error = yield* Effect.flip(read(input, { limits: { maxInputBytes: input.length - 1 } }))
        assert.deepStrictEqual(
          error._tag === "LimitExceeded" && [error.limit, error.max, error.actual],
          ["maxInputBytes", input.length - 1, input.length],
          name
        )
        const workbook = yield* read(input, { limits: { maxInputBytes: input.length } })
        assert.strictEqual(workbook.format, format, name)
        assert.strictEqual(workbook.limits.maxInputBytes, input.length, name)
        const error2 = yield* Effect.flip(Spreadsheet.detectFormat(input, { limits: { maxInputBytes: input.length - 1 } }))
        assert.strictEqual(error2._tag === "LimitExceeded" && error2.actual, input.length, name)
        assert.strictEqual(yield* Spreadsheet.detectFormat(input, { limits: { maxInputBytes: input.length } }), format, name)
      }
    }))

  it.effect("the size limit wins over a bad CSV option; a byte stream's size is checked while reading", () =>
    Effect.gen(function*() {
      const bytes = enc("a,b\n")
      const tooBig = { limits: { maxInputBytes: 1 } }
      const fromCsv = yield* Effect.flip(Csv.read(bytes, { ...tooBig, delimiter: "ab" }))
      assert.strictEqual(fromCsv._tag, "LimitExceeded")
      const fromRead = yield* Effect.flip(Spreadsheet.read(bytes, { ...tooBig, csv: { encoding: "no-such-encoding" } }))
      assert.strictEqual(fromRead._tag, "LimitExceeded")
      assert.strictEqual((yield* Effect.flip(Csv.read(bytes, { delimiter: "ab" })))._tag, "InvalidOption")
      assert.strictEqual((yield* Effect.flip(Spreadsheet.read(bytes, { csv: { encoding: "no-such-encoding" } })))._tag, "InvalidOption")
      // a stream has no size up front: the workbook opens, its rows fail
      const workbook = yield* Csv.read(Stream.make(bytes), tooBig)
      const error = yield* Effect.flip(Stream.runCollect(Workbook.rows(workbook, 0)))
      assert.strictEqual(error._tag === "LimitExceeded" && error.limit, "maxInputBytes")
      assert.strictEqual((yield* Effect.flip(Csv.read(Stream.make(bytes), { delimiter: "ab" })))._tag, "InvalidOption")
    }))

  it.effect("invalid limits, in the context or the call, die at every entry point, naming the limit", () =>
    Effect.gen(function*() {
      const bytes = yield* inputs
      const badContext = <A, E>(effect: Effect.Effect<A, E>) =>
        Effect.provideService(effect, SpreadsheetLimits.SpreadsheetLimits, { ...SpreadsheetLimits.defaults, maxSheets: NaN })
      for (const [name, format, read] of readers) {
        const input = bytes[format]
        assert.match(yield* dies(badContext(read(input))), /SpreadsheetLimits\.maxSheets must be a non-negative integer or Infinity, got NaN/, name)
        assert.match(yield* dies(read(input, { limits: { maxRows: -1 } })), /SpreadsheetLimits\.maxRows must be .*, got -1/, name)
        assert.match(yield* dies(badContext(Spreadsheet.detectFormat(input))), /maxSheets/, name)
        assert.match(yield* dies(Spreadsheet.detectFormat(input, { limits: { maxRows: 1.5 } })), /maxRows .*got 1\.5/, name)
      }
      assert.match(yield* dies(badContext(Csv.read(Stream.make(bytes.csv)))), /maxSheets/)
      assert.match(yield* dies(Csv.read(Stream.make(bytes.csv), { limits: { maxRows: -1 } })), /maxRows/)
    }))

  it.effect("the options are read when the effect runs: a re-run sees options changed after it was built", () =>
    Effect.gen(function*() {
      const bytes = yield* inputs
      for (const [name, format, read] of readers) {
        const input = bytes[format]
        const options: { limits?: SpreadsheetLimits.LimitOverrides } = {}
        const effect = read(input, options)
        assert.strictEqual((yield* effect).format, format, name)
        options.limits = { maxInputBytes: input.length - 1 }
        const error = yield* Effect.flip(effect)
        assert.strictEqual(error._tag === "LimitExceeded" && error.limit, "maxInputBytes", name)
        options.limits = { maxRows: 2 }
        assert.strictEqual((yield* effect).limits.maxRows, 2, name)
      }
    }))

  it.effect("the reader reads with the call's limits, on top of the context's", () =>
    Effect.gen(function*() {
      const bytes = yield* inputs
      for (const [name, format, read] of readers) {
        const workbook = yield* read(bytes[format], { limits: { maxRows: 2 } }).pipe(
          Effect.provideService(SpreadsheetLimits.SpreadsheetLimits, { ...SpreadsheetLimits.defaults, maxColumns: 7 })
        )
        assert.deepStrictEqual(workbook.limits, { ...SpreadsheetLimits.defaults, maxColumns: 7, maxRows: 2 }, name)
        const error = yield* Effect.flip(Stream.runCollect(Workbook.rows(workbook, 0)))
        assert.strictEqual(error._tag === "LimitExceeded" && error.limit, "maxRows", name)
      }
    }))
})

describe("Workbook.rows and sheetDetails: the call's overrides on the workbook's limits", () => {
  it.effect("overrides apply to that call only; undefined inherits", () =>
    Effect.gen(function*() {
      const bytes = yield* inputs
      for (const [name, format, read] of readers) {
        const workbook = yield* read(bytes[format], { limits: { maxRows: 2 } })
        const rows = yield* Stream.runCollect(Workbook.rows(workbook, 0, { limits: { maxRows: 3 } }))
        assert.deepStrictEqual(values(Array.from(rows)), [["1"], ["2"], ["3"]], name)
        // the override did not change the workbook's own limits
        assert.strictEqual(workbook.limits.maxRows, 2, name)
        for (const limits of [undefined, { maxRows: undefined }]) {
          const error = yield* Effect.flip(Stream.runCollect(Workbook.rows(workbook, 0, { limits })))
          assert.strictEqual(error._tag === "LimitExceeded" && error.limit, "maxRows", name)
        }
        yield* Workbook.sheetDetails(workbook, 0, { limits: { maxRows: 3 } })
      }
    }))

  it.effect("an invalid override dies when run, naming the first invalid limit in key order", () =>
    Effect.gen(function*() {
      const bytes = yield* inputs
      for (const [name, format, read] of readers) {
        const workbook = yield* read(bytes[format])
        // listed out of key order: maxRows comes before maxStringLength
        const limits = { maxStringLength: NaN, maxRows: -1 }
        const rows = Workbook.rows(workbook, 0, { limits }) // built, not run: no defect yet
        assert.match(yield* dies(Stream.runCollect(rows)), /SpreadsheetLimits\.maxRows must be a non-negative integer or Infinity, got -1/, name)
        const details = Workbook.sheetDetails(workbook, 0, { limits })
        assert.match(yield* dies(details), /SpreadsheetLimits\.maxRows must be .*, got -1/, name)
        assert.match(yield* dies(Stream.runCollect(Workbook.rows(workbook, 0, { limits: { maxColumns: "3" as unknown as number } }))), /maxColumns .*got string/, name)
      }
    }))

  it.effect("limits that no reader made are validated in full", () =>
    Effect.gen(function*() {
      const sheet: Workbook.SheetInfo = { name: "S", index: 0, visibility: "visible" }
      const handMade = Workbook.make({
        format: "csv",
        sheets: [sheet],
        limits: { ...SpreadsheetLimits.defaults, maxCellsPerSheet: NaN },
        rows: () => Stream.empty
      })
      assert.match(yield* dies(Stream.runCollect(Workbook.rows(handMade, 0))), /maxCellsPerSheet/)
      assert.match(yield* dies(Workbook.sheetDetails(handMade, 0, { limits: { maxRows: 1 } })), /maxCellsPerSheet/)
      // a reader's workbook with its limits replaced
      const read = yield* Csv.read(enc("1\n"))
      const copy: Workbook.Workbook = { ...read, limits: { ...read.limits, maxRows: NaN } }
      assert.match(yield* dies(Stream.runCollect(Workbook.rows(copy, 0))), /maxRows/)
      // valid hand-made limits read as before
      const valid = Workbook.make({ format: "csv", sheets: [sheet], limits: SpreadsheetLimits.defaults, rows: () => Stream.empty })
      assert.strictEqual((yield* Stream.runCollect(Workbook.rows(valid, 0, { limits: { maxRows: 1 } }))).length, 0)
    }))

  it.effect("a reader's workbook limits are trusted: changing them in place is not checked again", () =>
    Effect.gen(function*() {
      const bytes = yield* inputs
      for (const [name, format, read] of readers) {
        const workbook = yield* read(bytes[format])
        // the fields are readonly; a cast writes an invalid value anyway
        ;(workbook.limits as { maxRows: number }).maxRows = NaN
        // NaN disables the row check rather than making the read die
        const rows = yield* Stream.runCollect(Workbook.rows(workbook, 0))
        assert.strictEqual(rows.length, 3, name)
        assert.strictEqual((yield* Stream.runCollect(Workbook.rows(workbook, 0, { limits: { maxColumns: 5 } }))).length, 3, name)
      }
    }))

  it.effect("ValidLimits come only from validation", () =>
    Effect.gen(function*() {
      // @ts-expect-error plain Limits are not ValidLimits
      const plain: ValidLimits = SpreadsheetLimits.defaults
      // @ts-expect-error mergeValid takes ValidLimits as its base
      mergeValid(SpreadsheetLimits.defaults, undefined)
      assert.strictEqual(plain, SpreadsheetLimits.defaults)
      const valid = yield* resolveLimits()
      assert.strictEqual(mergeValid(valid, undefined), valid)
      const merged: SpreadsheetLimits.Limits = mergeValid(valid, { maxRows: 1, maxColumns: undefined })
      assert.deepStrictEqual(merged, { ...SpreadsheetLimits.defaults, maxRows: 1 })
      assert.strictEqual(valid.maxRows, SpreadsheetLimits.defaults.maxRows)
    }))
})

/**
 * The model, error and reader types are derived from their Schemas or
 * written as overloads; these checks pin them to the shapes the public API
 * documents, so a Schema edit that changes a public type fails here (at
 * `pnpm check` for the type checks).
 */
import { assert, describe, expectTypeOf, it } from "@effect/vitest"
import { Effect, Schema, Stream } from "effect"
import type * as Cell from "../src/Cell.ts"
import * as Csv from "../src/csv/index.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as SpreadsheetError from "../src/SpreadsheetError.ts"
import * as Workbook from "../src/Workbook.ts"

describe("types derived from their Schemas", () => {
  it("Cell is the union of the six cell kinds", () => {
    expectTypeOf<Cell.Cell>().toEqualTypeOf<
      Cell.EmptyCell | Cell.TextCell | Cell.NumberCell | Cell.BooleanCell | Cell.DateCell | Cell.ErrorCell
    >()
    expectTypeOf<Cell.TextCell>().toEqualTypeOf<{
      readonly _tag: "Text"
      readonly column: number
      readonly formula?: string
      readonly numFmt?: string
      readonly value: string
    }>()
  })

  it("Row has an index, a hidden flag and its cells", () => {
    expectTypeOf<Workbook.Row>().toEqualTypeOf<{
      readonly index: number
      readonly hidden: boolean
      readonly cells: ReadonlyArray<Cell.Cell>
    }>()
  })

  it("SpreadsheetError is the union of the six error classes", () => {
    expectTypeOf<SpreadsheetError.SpreadsheetError>().toEqualTypeOf<
      | SpreadsheetError.UnsupportedFormat
      | SpreadsheetError.MalformedFile
      | SpreadsheetError.LimitExceeded
      | SpreadsheetError.Encrypted
      | SpreadsheetError.SheetNotFound
      | SpreadsheetError.InvalidOption
    >()
  })
})

describe("isSpreadsheetError", () => {
  it("passes an instance of each class in the union schema, and nothing else with the same tag", () => {
    const errors = [
      new SpreadsheetError.UnsupportedFormat({ format: "pdf", reason: "x" }),
      new SpreadsheetError.MalformedFile({ format: "zip", reason: "x" }),
      new SpreadsheetError.LimitExceeded({ limit: "maxRows", max: 1, actual: 2 }),
      new SpreadsheetError.Encrypted({ format: "xlsx" }),
      new SpreadsheetError.SheetNotFound({ sheet: 0 }),
      new SpreadsheetError.InvalidOption({ option: "header", reason: "x" })
    ]
    assert.strictEqual(SpreadsheetError.SpreadsheetError.members.length, errors.length)
    for (const e of errors) assert.isTrue(SpreadsheetError.isSpreadsheetError(e))
    // another library's error with the same tag and fields is not ours
    class InvalidOption extends Schema.TaggedError<InvalidOption>()("InvalidOption", {
      option: Schema.String,
      reason: Schema.String
    }) {}
    assert.isFalse(SpreadsheetError.isSpreadsheetError(new InvalidOption({ option: "a", reason: "b" })))
    assert.isFalse(SpreadsheetError.isSpreadsheetError({ _tag: "InvalidOption", option: "a", reason: "b" }))
    assert.isFalse(SpreadsheetError.isSpreadsheetError(null))
  })
})

describe("dual-input readers", () => {
  it("toRecords gives a stream for a stream and an Effect of an array for an iterable", () => {
    const rows: Array<Workbook.Row> = [{ index: 0, hidden: false, cells: [] }]
    const fromStream = Spreadsheet.toRecords(
      Stream.fromIterable(rows).pipe(Stream.mapError((e: "upstream") => e)),
      { header: "first-row", defval: null }
    )
    expectTypeOf(fromStream).toEqualTypeOf<
      Stream.Stream<
        Record<string, Cell.Value | null>,
        "upstream" | SpreadsheetError.LimitExceeded | SpreadsheetError.InvalidOption
      >
    >()
    const fromArray = Spreadsheet.toRecords(rows, { header: ["a"] })
    expectTypeOf(fromArray).toEqualTypeOf<
      Effect.Effect<Array<Record<string, Cell.Value>>, SpreadsheetError.LimitExceeded | SpreadsheetError.InvalidOption>
    >()
    assert.isTrue(Stream.isStream(fromStream))
    assert.isTrue(Effect.isEffect(fromArray))
  })

  it("Csv.rows keeps a byte stream's error and context", () => {
    const fromBytes = Csv.rows(new Uint8Array())
    expectTypeOf(fromBytes).toEqualTypeOf<Stream.Stream<Workbook.Row, SpreadsheetError.SpreadsheetError>>()
    const fromStream = Csv.rows(Stream.succeed(new Uint8Array()).pipe(Stream.mapError((e: "upstream") => e)))
    expectTypeOf(fromStream).toEqualTypeOf<
      Stream.Stream<Workbook.Row, "upstream" | SpreadsheetError.SpreadsheetError>
    >()
  })

  it("the workbook's source keys are unique symbols, shared across copies of the module", () => {
    expectTypeOf<typeof Workbook.RowsSource>().not.toEqualTypeOf<symbol>()
    expectTypeOf<typeof Workbook.DetailsSource>().not.toEqualTypeOf<symbol>()
    assert.strictEqual(Workbook.RowsSource, Symbol.for("effect-spreadsheet/Workbook/rows"))
    assert.strictEqual(Workbook.DetailsSource, Symbol.for("effect-spreadsheet/Workbook/details"))
  })
})

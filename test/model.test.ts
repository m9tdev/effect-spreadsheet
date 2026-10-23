import { assert, describe, it } from "@effect/vitest"
import { Effect, Layer, Schema, Stream } from "effect"
import * as Cell from "../src/Cell.ts"
import * as SpreadsheetError from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"

describe("Cell", () => {
  it("constructors build plain tagged objects that the schema accepts", () => {
    const cells: Array<Cell.Cell> = [
      Cell.empty(0, { numFmt: "0.00" }),
      Cell.text(1, "a"),
      Cell.number(2, 1.5, { formula: "A1*2" }),
      Cell.boolean(3, true),
      Cell.date(4, new Date(Date.UTC(2024, 0, 2))),
      Cell.error(5, "#N/A")
    ]
    for (const cell of cells) assert.deepStrictEqual(Schema.decodeUnknownSync(Cell.Cell)(cell), cell)
    assert.deepStrictEqual(cells.map((c) => c._tag), ["Empty", "Text", "Number", "Boolean", "Date", "Error"])
    assert.isFalse("formula" in cells[1]!)
    assert.strictEqual(cells[2]!.formula, "A1*2")
    assert.strictEqual(Cell.value(cells[0]!), undefined)
    assert.strictEqual(Cell.value(cells[5]!), "#N/A")
    assert.isTrue(Cell.isEmpty(cells[0]!))
  })

  it("schema rejects bad cells", () => {
    assert.throws(() => Schema.decodeUnknownSync(Cell.Cell)({ _tag: "Number", column: 0, value: "x" }))
    assert.throws(() => Schema.decodeUnknownSync(Cell.Cell)({ _tag: "Text", column: 0.5, value: "x" }))
  })
})

describe("SpreadsheetError", () => {
  it.effect("errors are tagged, yieldable and in the union schema", () =>
    Effect.gen(function*() {
      const errors = [
        new SpreadsheetError.UnsupportedFormat({ format: "xls", reason: "later" }),
        new SpreadsheetError.MalformedFile({ format: "zip", part: "xl/workbook.xml", offset: 12, reason: "bad" }),
        new SpreadsheetError.LimitExceeded({ limit: "maxRows", max: 1, actual: 2 }),
        new SpreadsheetError.Encrypted({ format: "xlsx" }),
        new SpreadsheetError.SheetNotFound({ sheet: "x" })
      ]
      for (const e of errors) {
        const flipped = yield* Effect.flip(Effect.fail(e))
        assert.strictEqual(flipped, e)
        assert.isTrue(SpreadsheetError.isSpreadsheetError(e))
        assert.isTrue(Schema.is(SpreadsheetError.SpreadsheetError)(e))
        assert.isTrue(e.message.length > 0)
      }
      assert.strictEqual(errors[1]!.message, "Malformed zip file in xl/workbook.xml at offset 12: bad")
      assert.isFalse(SpreadsheetError.isSpreadsheetError(new Error("x")))
      const caught = yield* Effect.fail(errors[2]!).pipe(Effect.catchTag("LimitExceeded", (e) => Effect.succeed(e.limit)))
      assert.strictEqual(caught, "maxRows")
    }))
})

describe("SpreadsheetLimits", () => {
  it.effect("defaults without a layer", () =>
    Effect.gen(function*() {
      const limits = yield* SpreadsheetLimits.SpreadsheetLimits
      assert.deepStrictEqual(limits, SpreadsheetLimits.defaults)
      assert.strictEqual(limits.maxColumns, 16_384)
      assert.strictEqual(limits.maxRows, 1_048_576)
    }))

  it.effect("layer overrides some keys, call overrides win", () =>
    Effect.gen(function*() {
      const fromLayer = yield* SpreadsheetLimits.resolve()
      assert.strictEqual(fromLayer.maxRows, 10)
      assert.strictEqual(fromLayer.maxColumns, 16_384)
      const perCall = yield* SpreadsheetLimits.resolve({ maxRows: 5, maxColumns: undefined })
      assert.strictEqual(perCall.maxRows, 5)
      assert.strictEqual(perCall.maxColumns, 16_384)
    }).pipe(Effect.provide(SpreadsheetLimits.layer({ maxRows: 10 }))))

  it("mergePartial keeps defined values", () => {
    assert.deepStrictEqual(SpreadsheetLimits.mergePartial({ maxRows: 1 }, { maxRows: undefined, maxColumns: 2 }), {
      maxRows: 1,
      maxColumns: 2
    })
    assert.deepStrictEqual(SpreadsheetLimits.mergePartial(undefined, { maxRows: 3 }), { maxRows: 3 })
  })

  it.effect("Layer.succeed with a full object works too", () =>
    Effect.gen(function*() {
      const limits = yield* SpreadsheetLimits.SpreadsheetLimits
      assert.strictEqual(limits.maxSheets, 1)
    }).pipe(Effect.provide(Layer.succeed(SpreadsheetLimits.SpreadsheetLimits, { ...SpreadsheetLimits.defaults, maxSheets: 1 }))))
})

describe("Workbook", () => {
  const wb = Workbook.make({
    format: "test",
    sheets: [
      { name: "Data", index: 0, visibility: "visible" },
      { name: "Hidden", index: 1, visibility: "hidden", dimension: { firstRow: 0, firstColumn: 0, lastRow: 1048575, lastColumn: 16383 } }
    ],
    rows: (sheet) => Stream.make({ index: 0, hidden: sheet.index === 1, cells: [Cell.text(0, sheet.name)] })
  })

  it.effect("streams rows by name (case-insensitive fallback) or index", () =>
    Effect.gen(function*() {
      const byName = yield* Stream.runCollect(Workbook.rows(wb, "Data"))
      assert.strictEqual(Cell.value(Array.from(byName)[0]!.cells[0]!), "Data")
      const byLower = yield* Stream.runCollect(Workbook.rows(wb, "hidden"))
      assert.isTrue(Array.from(byLower)[0]!.hidden)
      const byIndex = yield* Stream.runCollect(Workbook.rows(wb, 1))
      assert.strictEqual(Cell.value(Array.from(byIndex)[0]!.cells[0]!), "Hidden")
      assert.strictEqual((yield* Workbook.sheet(wb, 0)).name, "Data")
    }))

  it.effect("missing sheet fails with SheetNotFound", () =>
    Effect.gen(function*() {
      const e1 = yield* Effect.flip(Stream.runCollect(Workbook.rows(wb, "Nope")))
      assert.strictEqual(e1._tag, "SheetNotFound")
      const e2 = yield* Effect.flip(Workbook.sheet(wb, 7))
      assert.deepStrictEqual(e2.sheet, 7)
    }))

  it("SheetInfo and Row schemas validate", () => {
    for (const s of wb.sheets) Schema.decodeUnknownSync(Workbook.SheetInfo)(s)
    Schema.decodeUnknownSync(Workbook.Row)({ index: 3, hidden: true, cells: [Cell.number(16383, 1)] })
    assert.throws(() => Schema.decodeUnknownSync(Workbook.SheetInfo)({ name: "x", index: 0, visibility: "gone" }))
  })
})

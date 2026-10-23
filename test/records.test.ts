import { assert, describe, it } from "@effect/vitest"
import { Effect, Stream } from "effect"
import * as Cell from "../src/Cell.ts"
import * as Csv from "../src/csv/index.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import type { Row } from "../src/Workbook.ts"
import { csvRows, enc } from "./helpers/run.ts"

const rowsOf = (text: string) => Effect.runPromise(csvRows(text, { inferTypes: true }))

describe("toRecords", () => {
  it("first-row header", async () => {
    const rows = await rowsOf("name,age\nann,31\nbob,\n")
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: "first-row" }), [
      { name: "ann", age: 31 },
      { name: "bob" }
    ])
  })

  it("defval fills missing header columns; blank rows skipped unless blankRows", async () => {
    const rows = await rowsOf("a,b\n1\n\n,2\n")
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: "first-row", defval: null }), [
      { a: 1, b: null },
      { a: null, b: 2 }
    ])
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: "first-row", blankRows: true }), [{ a: 1 }, {}, { b: 2 }])
  })

  it("duplicate and empty header names are made unique like SheetJS", async () => {
    const rows = await rowsOf("a,a,,a,,a_1\n1,2,3,4,5,6\n")
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: "first-row" }), [
      { a: 1, a_1: 2, __EMPTY: 3, a_2: 4, __EMPTY_1: 5, a_1_1: 6 }
    ])
  })

  it("columns past the header get __EMPTY names", async () => {
    const rows = await rowsOf("a\n1,2,3\n")
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: "first-row" }), [{ a: 1, __EMPTY: 2, __EMPTY_1: 3 }])
  })

  it("header by row index skips earlier rows", async () => {
    const rows = await rowsOf("title\n\nx,y\n1,2\n")
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: 2 }), [{ x: 1, y: 2 }])
  })

  it("header row missing from sparse rows: columns get __EMPTY names", () => {
    const rows: Array<Row> = [{ index: 5, hidden: false, cells: [Cell.text(1, "v")] }]
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: 2 }), [{ __EMPTY_1: "v" }])
  })

  it("explicit header names; extra columns dropped", async () => {
    const rows = await rowsOf("1,2,3\n4,5\n")
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: ["x", "y"] }), [{ x: 1, y: 2 }, { x: 4, y: 5 }])
  })

  it("__proto__ header becomes an own key, not the prototype", async () => {
    const rows = await rowsOf("__proto__,b\n1,2\n")
    const [record] = Spreadsheet.toRecords(rows, { header: "first-row" })
    assert.strictEqual(Object.getPrototypeOf(record), Object.prototype)
    assert.deepStrictEqual(Object.keys(record!), ["__proto__", "b"])
    assert.strictEqual(Object.getOwnPropertyDescriptor(record, "__proto__")?.value, 1)
  })

  it("dates in header become ISO text; values are cell values", () => {
    const d = new Date(Date.UTC(2024, 0, 2))
    const rows: Array<Row> = [
      { index: 0, hidden: false, cells: [Cell.date(0, d), Cell.number(1, 1)] },
      { index: 1, hidden: true, cells: [Cell.boolean(0, true), Cell.error(1, "#N/A"), Cell.empty(2)] }
    ]
    assert.deepStrictEqual(Spreadsheet.toRecords(rows, { header: "first-row" }), [
      { "2024-01-02T00:00:00.000Z": true, "1": "#N/A" }
    ])
  })

  it.effect("works lazily on streams", () =>
    Effect.gen(function*() {
      const stream = Spreadsheet.toRecords(Csv.rows(enc("k,v\na,1\nb,2\n"), { inferTypes: true }), { header: "first-row" })
      const records = yield* Stream.runCollect(stream)
      assert.deepStrictEqual(Array.from(records), [{ k: "a", v: 1 }, { k: "b", v: 2 }])
      // re-runnable: the header state is per run
      const again = yield* Stream.runCollect(stream)
      assert.deepStrictEqual(Array.from(again), Array.from(records))
    }))
})

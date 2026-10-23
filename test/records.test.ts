import { assert, describe, expect, expectTypeOf, it } from "@effect/vitest"
import type { LimitExceeded } from "../src/SpreadsheetError.ts"
import { Cause, Effect, Stream } from "effect"
import * as Cell from "../src/Cell.ts"
import * as Csv from "../src/csv/index.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import type { Row } from "../src/Workbook.ts"
import { csvRows, enc } from "./helpers/run.ts"
import { countYields } from "./helpers/yields.ts"

const rowsOf = (text: string) => Effect.runPromise(csvRows(text, { inferTypes: true }))
const recs = <D = never>(rows: Iterable<Row>, options: Spreadsheet.RecordsOptions<D>) =>
  Effect.runSync(Spreadsheet.toRecords(rows, options))

describe("toRecords", () => {
  it("first-row header", async () => {
    const rows = await rowsOf("name,age\nann,31\nbob,\n")
    assert.deepStrictEqual(recs(rows, { header: "first-row" }), [
      { name: "ann", age: 31 },
      { name: "bob" }
    ])
  })

  it("defval fills missing header columns; blank rows skipped unless blankRows", async () => {
    const rows = await rowsOf("a,b\n1\n\n,2\n")
    assert.deepStrictEqual(recs(rows, { header: "first-row", defval: null }), [
      { a: 1, b: null },
      { a: null, b: 2 }
    ])
    assert.deepStrictEqual(recs(rows, { header: "first-row", blankRows: true }), [{ a: 1 }, {}, { b: 2 }])
  })

  it("duplicate and empty header names are made unique like SheetJS", async () => {
    const rows = await rowsOf("a,a,,a,,a_1\n1,2,3,4,5,6\n")
    assert.deepStrictEqual(recs(rows, { header: "first-row" }), [
      { a: 1, a_1: 2, __EMPTY: 3, a_2: 4, __EMPTY_1: 5, a_1_1: 6 }
    ])
  })

  it("unnamed columns get __EMPTY, then __EMPTY_1, __EMPTY_2, in order of first data", async () => {
    const rows = await rowsOf("a\n,,,1\n,2,3\n")
    assert.deepStrictEqual(recs(rows, { header: "first-row" }), [{ __EMPTY: 1 }, { __EMPTY_1: 2, __EMPTY_2: 3 }])
  })

  it("columns past the header get __EMPTY names", async () => {
    const rows = await rowsOf("a\n1,2,3\n")
    assert.deepStrictEqual(recs(rows, { header: "first-row" }), [{ a: 1, __EMPTY: 2, __EMPTY_1: 3 }])
  })

  it("header by row index skips earlier rows", async () => {
    const rows = await rowsOf("title\n\nx,y\n1,2\n")
    assert.deepStrictEqual(recs(rows, { header: 2 }), [{ x: 1, y: 2 }])
  })

  it("header row missing from sparse rows: columns get __EMPTY names", () => {
    const rows: Array<Row> = [{ index: 5, hidden: false, cells: [Cell.text(1, "v")] }]
    assert.deepStrictEqual(recs(rows, { header: 2 }), [{ __EMPTY: "v" }])
  })

  it("explicit header names; extra columns dropped", async () => {
    const rows = await rowsOf("1,2,3\n4,5\n")
    assert.deepStrictEqual(recs(rows, { header: ["x", "y"] }), [{ x: 1, y: 2 }, { x: 4, y: 5 }])
  })

  it("explicit header with a hole, undefined or null drops that column", async () => {
    const rows = await rowsOf("1,2,3\n")
    const sparse: Array<string> = new Array(3)
    sparse[0] = "x"
    sparse[2] = "y"
    assert.deepStrictEqual(recs(rows, { header: sparse }), [{ x: 1, y: 3 }])
    assert.deepStrictEqual(recs(rows, { header: ["x", undefined as unknown as string, "y"] }), [{ x: 1, y: 3 }])
    assert.deepStrictEqual(recs(rows, { header: ["x", null as unknown as string, "y"] }), [{ x: 1, y: 3 }])
    // dropped columns get no defval either
    assert.deepStrictEqual(recs(await rowsOf("1\n"), { header: sparse, defval: 0 }), [{ x: 1, y: 0 }])
  })

  it("an explicit \"\" entry is like a blank header cell: __EMPTY key, no defval fill", async () => {
    assert.deepStrictEqual(recs(await rowsOf("1\n"), { header: ["x", "", "y"], defval: 0 }), [{ x: 1, y: 0 }])
    assert.deepStrictEqual(recs(await rowsOf("1,2\n"), { header: ["x", "", "y"], defval: 0 }), [{ x: 1, __EMPTY: 2, y: 0 }])
  })

  it("explicit header names are de-duplicated too, so no column is lost", async () => {
    const rows = await rowsOf("1,2,3,4\n")
    assert.deepStrictEqual(recs(rows, { header: ["x", "x", "", "x_1"] }), [{ x: 1, x_1: 2, __EMPTY: 3, x_1_1: 4 }])
  })

  it("a TRUE/FALSE header with inferTypes keeps the keys TRUE and FALSE", async () => {
    const rows = await rowsOf("TRUE,FALSE,x\n1,2,3\n")
    assert.strictEqual(rows[0]!.cells[0]!._tag, "Boolean")
    assert.deepStrictEqual(recs(rows, { header: "first-row" }), [{ TRUE: 1, FALSE: 2, x: 3 }])
  })

  it("header cells of every type become names; an Empty header cell is unnamed", () => {
    const rows: Array<Row> = [
      { index: 0, hidden: false, cells: [Cell.boolean(0, true), Cell.error(1, "#N/A"), Cell.empty(2), Cell.number(3, 1.5)] },
      { index: 1, hidden: false, cells: [Cell.text(0, "a"), Cell.text(1, "b"), Cell.text(2, "c"), Cell.text(3, "d")] }
    ]
    assert.deepStrictEqual(recs(rows, { header: "first-row", defval: null }), [
      { TRUE: "a", "#N/A": "b", "1.5": "d", __EMPTY: "c" }
    ])
  })

  it("__proto__ header becomes an own key, not the prototype", async () => {
    const rows = await rowsOf("__proto__,b\n1,2\n")
    const [record] = recs(rows, { header: "first-row" })
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
    assert.deepStrictEqual(recs(rows, { header: "first-row" }), [
      { "2024-01-02T00:00:00.000Z": true, "1": "#N/A" }
    ])
  })

  it("defval fills only header-named columns; far cells don't widen every record (B1 regression)", async () => {
    // header `a`, one row with a cell at column 16383, then 5000 rows of `y`
    const text = "a\n" + ",".repeat(16383) + "far\n" + "y\n".repeat(5000)
    expect(text.length).toBeLessThan(30_000)
    const rows = await Effect.runPromise(csvRows(text))
    // The key count below is the cost check (it was 81.9M before the fix);
    // the clock only has a generous ceiling against a gross blowup.
    const start = performance.now()
    const records = recs(rows, { header: "first-row", defval: null })
    expect(performance.now() - start).toBeLessThan(10_000)
    expect(records.length).toBe(5001)
    expect(records[0]).toEqual({ a: null, __EMPTY: "far" })
    expect(records[1]).toEqual({ a: "y" })
    const keys = records.reduce((n, r) => n + Object.keys(r).length, 0)
    expect(keys).toBe(5002)
  })

  it("gap columns inside the header get no defval and are named when data appears", async () => {
    const rows = await rowsOf("a,,c\n1\n,2\n")
    assert.deepStrictEqual(recs(rows, { header: "first-row", defval: 0 }), [
      { a: 1, c: 0 },
      { a: 0, c: 0, __EMPTY: 2 }
    ])
  })

  it.effect("cells count against maxCellsPerSheet, defval fills against maxRecordFields", () =>
    Effect.gen(function*() {
      const rows = yield* csvRows("a,b,c\n1\n2\n3\n")
      // 3 records x (1 value in "a" + 2 fills for "b", "c"): 6 fills, 3 cells
      const ok = yield* Spreadsheet.toRecords(rows, {
        header: "first-row",
        defval: null,
        limits: { maxCellsPerSheet: 3, maxRecordFields: 6 }
      })
      assert.strictEqual(ok.length, 3)
      assert.deepStrictEqual(ok[0], { a: "1", b: null, c: null })
      const fills = yield* Effect.flip(
        Spreadsheet.toRecords(rows, { header: "first-row", defval: null, limits: { maxRecordFields: 5 } })
      )
      assert.deepStrictEqual([fills._tag, fills.limit, fills.max, fills.actual], ["LimitExceeded", "maxRecordFields", 5, 6])
      const cells = yield* Effect.flip(
        Spreadsheet.toRecords(rows, { header: "first-row", limits: { maxCellsPerSheet: 2 } })
      )
      assert.deepStrictEqual([cells.limit, cells.max], ["maxCellsPerSheet", 2])
      // the stream variant fails the same way, and the layer applies
      const streamError = yield* Effect.flip(
        Stream.runCollect(Spreadsheet.toRecords(Stream.fromIterable(rows), { header: "first-row", defval: null }))
      ).pipe(Effect.provide(SpreadsheetLimits.layer({ maxRecordFields: 5 })))
      assert.strictEqual(streamError._tag === "LimitExceeded" && streamError.limit, "maxRecordFields")
    }))

  it.effect("a dense file with defval needs no fills: 60 000 x 20 values stay under maxRecordFields", () =>
    Effect.gen(function*() {
      const header = Array.from({ length: 20 }, (_, i) => `h${i}`).join(",")
      const row = Array.from({ length: 20 }, (_, i) => `v${i}`).join(",")
      const bytes = enc(header + "\n" + (row + "\n").repeat(60_000))
      const count = yield* Stream.runCount(Spreadsheet.toRecords(Csv.rows(bytes), { header: "first-row", defval: "" }))
      assert.strictEqual(count, 60_000)
      // Empty cells and missing columns are fills: 1 value + 1 Empty + 1 missing = 2 fills per record
      const sparse: Array<Row> = [
        { index: 0, hidden: false, cells: [Cell.text(0, "a"), Cell.text(1, "b"), Cell.text(2, "c")] },
        { index: 1, hidden: false, cells: [Cell.text(0, "x"), Cell.empty(1)] },
        { index: 2, hidden: false, cells: [Cell.text(0, "y"), Cell.empty(1)] }
      ]
      const ok = yield* Spreadsheet.toRecords(sparse, { header: "first-row", defval: 0, limits: { maxRecordFields: 4 } })
      assert.deepStrictEqual(ok, [{ a: "x", b: 0, c: 0 }, { a: "y", b: 0, c: 0 }])
      const over = yield* Effect.flip(
        Spreadsheet.toRecords(sparse, { header: "first-row", defval: 0, limits: { maxRecordFields: 3 } })
      )
      assert.strictEqual(over.limit, "maxRecordFields")
    }), 30_000)

  it("many rows that give no record (blank, before the header) finish fast and keep yielding", async () => {
    // ~600 KB: the reader hands over ~10 pieces, and the stream yields between them
    const bytes = enc("\n".repeat(300_000) + "h\n" + "\n".repeat(300_000) + "v\n")
    const { value: records, yields, ms } = await countYields(() =>
      Effect.runPromise(Stream.runCollect(Spreadsheet.toRecords(Csv.rows(bytes), { header: 300_000 })))
    )
    expect(Array.from(records)).toEqual([{ h: "v" }])
    expect(yields).toBeGreaterThanOrEqual(5)
    expect(ms).toBeLessThan(10_000)
  }, 30_000)

  it.effect("Empty cells count against maxCellsPerSheet", () =>
    Effect.gen(function*() {
      const rows: Array<Row> = [
        { index: 0, hidden: false, cells: [Cell.text(0, "h")] },
        { index: 1, hidden: false, cells: [Cell.empty(0), Cell.empty(1)] },
        { index: 2, hidden: false, cells: [Cell.empty(0)] }
      ]
      const error = yield* Effect.flip(Spreadsheet.toRecords(rows, { header: "first-row", limits: { maxCellsPerSheet: 2 } }))
      assert.deepStrictEqual([error.limit, error.actual], ["maxCellsPerSheet", 3])
    }))

  it.effect("records built before a limit hit in the same chunk are emitted first", () =>
    Effect.gen(function*() {
      const got: Array<unknown> = []
      const stream = Spreadsheet.toRecords(Csv.rows(enc("h\n1\n2\n3\n4"), { inferTypes: true }), {
        header: "first-row",
        limits: { maxCellsPerSheet: 2 }
      })
      const error = yield* Effect.flip(Stream.runForEach(stream, (r) => Effect.sync(() => got.push(r))))
      assert.deepStrictEqual(got, [{ h: 1 }, { h: 2 }])
      assert.strictEqual(error._tag === "LimitExceeded" && error.limit, "maxCellsPerSheet")
    }))

  it("B1': a 16384-column header times 1300 rows with defval stays bounded and yielding", async () => {
    const header = Array.from({ length: 16384 }, (_, i) => `c${i}`).join(",")
    const bytes = enc(header + "\n" + "y\n".repeat(1300))
    expect(bytes.length).toBeLessThan(110_000)
    // Cost: the fill budget stops it after 64 records. Responsiveness: the
    // number of event-loop turns, not a wall-clock gap, so load doesn't
    // matter. Records come in batches of ~64 Ki keys (4 records here) with a
    // yield before each: ~19 turns for 64 records, ~330 for 1300; one big
    // synchronous batch gives 2-3. The longest stretch is bounded by the
    // records emitted between two turns: 4 per batch here, at most 8 allowed
    // (one oversized batch anywhere in the run fails it). The clock only has
    // generous ceilings.
    let emitted = 0
    const first = await countYields(
      () =>
        Effect.runPromise(Effect.exit(Stream.runForEach(
          Spreadsheet.toRecords(Csv.rows(bytes), { header: "first-row", defval: null }),
          () => Effect.sync(() => emitted++)
        ))),
      () => emitted
    )
    // 1 048 576 fills / 16383 per record: 64 records, then maxRecordFields
    expect(first.value._tag).toBe("Failure")
    const error = first.value._tag === "Failure" ? Cause.squash(first.value.cause) : undefined
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxRecordFields", max: 1_048_576 })
    expect(emitted).toBe(64)
    expect(first.yields).toBeGreaterThanOrEqual(64 / 8)
    expect(first.maxPerTurn).toBeLessThanOrEqual(8)
    expect(first.ms).toBeLessThan(10_000)
    // with the fill budget raised, all 1300 records come out, still yielding
    let all = 0
    const run = await countYields(
      () =>
        Effect.runPromise(Stream.runForEach(
          Spreadsheet.toRecords(Csv.rows(bytes), { header: "first-row", defval: null, limits: { maxRecordFields: Infinity } }),
          () => Effect.sync(() => all++)
        )),
      () => all
    )
    expect(all).toBe(1300)
    expect(run.yields).toBeGreaterThanOrEqual(1300 / 8)
    expect(run.maxPerTurn).toBeLessThanOrEqual(8)
  }, 60_000)

  it("record values are typed: cell values, plus the defval type", () => {
    const rows: Array<Row> = [{ index: 0, hidden: false, cells: [Cell.text(0, "a")] }]
    expectTypeOf(Spreadsheet.toRecords(rows, { header: "first-row" })).toEqualTypeOf<
      Effect.Effect<Array<Record<string, Cell.Value>>, LimitExceeded>
    >()
    expectTypeOf(Spreadsheet.toRecords(rows, { header: "first-row", defval: null })).toEqualTypeOf<
      Effect.Effect<Array<Record<string, Cell.Value | null>>, LimitExceeded>
    >()
    expectTypeOf(Spreadsheet.toRecords(Stream.fromIterable(rows), { header: "first-row", defval: "" })).toEqualTypeOf<
      Stream.Stream<Record<string, Cell.Value>, LimitExceeded>
    >()
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

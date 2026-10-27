import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import { NO_STYLES } from "../src/internal/xlsx/parts.ts"
import { SheetRows } from "../src/internal/xlsx/sheet.ts"
import { XmlTokenizer } from "../src/internal/xml/tokenizer.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { failureOf } from "./helpers/run.ts"
import { xlsx } from "./helpers/xlsx.ts"

const readExit = async (data: string, limits?: LimitOverrides) => {
  const bytes = await xlsx({ sheets: [{ name: "S", data }] })
  return Effect.runPromise(Effect.exit(Effect.gen(function*() {
    const wb = yield* Xlsx.read(bytes, { limits })
    return Array.from(yield* Stream.runCollect(Workbook.rows(wb, 0)))
  })))
}

const read = async (data: string, limits?: LimitOverrides) => {
  const exit = await readExit(data, limits)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value.map((r) => [r.index, r.cells.map((c) => [c.column, c._tag, Cell.value(c) ?? null])])
}

/** A SheetRows fed by a tokenizer, one piece of XML at a time. */
const driver = (maxColumns: number) => {
  const limits = { ...SpreadsheetLimits.defaults, maxColumns }
  const rows = new SheetRows({ part: "xl/worksheets/sheet1.xml", strings: undefined, styles: NO_STYLES, date1904: false, limits })
  const tokenizer = new XmlTokenizer(rows, limits, "xl/worksheets/sheet1.xml")
  return {
    rows,
    feed: (text: string) => expect(tokenizer.feed(text)).toBeUndefined(),
    end: () => expect(tokenizer.end()).toBeUndefined()
  }
}

describe("xlsx sheet: repeated columns", () => {
  it("a row's repeated columns don't pile up: past 2 x maxColumns cells it is compacted, the last cell of a column kept", () => {
    const { end, feed, rows } = driver(8)
    feed(`<worksheet><sheetData><row r="1">`)
    let most = 0
    // 1 000 passes over columns H..A (unsorted): 8 000 cells for 8 columns
    for (let k = 0; k < 8000; k++) {
      feed(`<c r="${"HGFEDCBA"[k % 8]}1"><v>${k}</v></c>`)
      most = Math.max(most, rows.heldCells)
    }
    expect(most).toBeLessThanOrEqual(2 * 8 + 1)
    feed(`</row></sheetData></worksheet>`)
    end()
    expect(rows.out).toHaveLength(1)
    expect(rows.out[0]!.cells.map((c) => [c.column, Cell.value(c)])).toEqual(
      Array.from({ length: 8 }, (_, c) => [c, 8000 - 8 + 7 - c])
    )
  })

  it("after a compaction, columns without r follow the previous cell, and later cells are still ordered", () => {
    const { end, feed, rows } = driver(4)
    // B A D C B A D C B: 9 cells > 2 x 4, compacted at the 9th (B1 = 8), leaving A..D
    feed(`<worksheet><sheetData><row r="1">`)
    for (let k = 0; k < 9; k++) feed(`<c r="${"DCBA"[(k + 2) % 4]}1"><v>${k}</v></c>`)
    expect(rows.heldCells).toBe(4)
    // no r: the column after the previous cell (B), i.e. C, which is before
    // the last held cell (D), so the row is unsorted again and C1 = 100 wins
    feed(`<c><v>100</v></c></row></sheetData></worksheet>`)
    end()
    expect(rows.out[0]!.cells.map((c) => [c.column, Cell.value(c)])).toEqual([[0, 5], [1, 8], [2, 100], [3, 6]])
  })

  it("duplicate columns: the last cell of a column wins, as before compaction existed", async () => {
    const patterns: ReadonlyArray<string> = [
      `<row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c><c r="A1"><v>3</v></c></row>`,
      `<row r="1"><c r="C1"><v>1</v></c><c r="A1"><v>2</v></c><c r="C1" t="b"><v>1</v></c><c><v>4</v></c><c r="A1"/></row>`,
      `<row r="2"><c r="B2"><v>1</v></c><c r="B2"><v>2</v></c></row><row r="3"><c r="B3"><v>3</v></c><c r="A3"><v>4</v></c></row>`
    ]
    const expected = [
      [[0, [[0, "Number", 3]]]],
      [[0, [[0, "Empty", null], [2, "Boolean", true], [3, "Number", 4]]]],
      [[1, [[1, "Number", 2]]], [2, [[0, "Number", 4], [1, "Number", 3]]]]
    ]
    for (const [k, data] of patterns.entries()) {
      expect(await read(data)).toEqual(expected[k])
      // compacting while the row is read (maxColumns 4: past 8 cells) gives the same result
      const many = data.replace(/<row r="(\d)">(.*?)<\/row>/g, (_, r, cells) => `<row r="${r}">${cells.repeat(5)}</row>`)
      expect(await read(many, { maxColumns: 4 })).toEqual(expected[k])
    }
  })

  it("duplicate cells still count against maxCellsPerSheet", async () => {
    const error = failureOf(await readExit(`<row r="1">${`<c r="A1"/>`.repeat(3)}</row>`, { maxCellsPerSheet: 2 }))
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxCellsPerSheet", max: 2, actual: 3 })
  })
})

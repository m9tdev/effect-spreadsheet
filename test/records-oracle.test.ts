/**
 * toRecords vs the SheetJS oracle (sheet_to_json, object mode, raw values;
 * skipped when SheetJS is absent, required in CI). Pins the intended
 * differences listed in docs/records-oracle-differences.md, and the cases
 * where both agree.
 */
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Csv from "../src/csv/index.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { enc } from "./helpers/run.ts"
import { xlsx as buildXlsx } from "./helpers/xlsx.ts"
import { load } from "./oracle/sheetjs.ts"

const xlsx = await load()

type Options = { header: "first-row" | ReadonlyArray<string>; defval?: unknown }

const ours = async (text: string, options: Options) =>
  Array.from(
    await Effect.runPromise(Stream.runCollect(Spreadsheet.toRecords(Csv.rows(enc(text), { inferTypes: true }), options)))
  )

const theirs = (text: string, options: Options) => {
  const wb = xlsx!.read(enc(text), { type: "buffer" })
  const o: Record<string, unknown> = { raw: true }
  if (Array.isArray(options.header)) o["header"] = options.header
  if ("defval" in options) o["defval"] = options.defval
  return xlsx!.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]!], o)
}

// Needs no oracle, so it runs even when SheetJS is absent.
describe("toRecords docs examples", () => {
  it("difference 5 lists the keys in JavaScript order, as the docs show", async () => {
    const [record] = await ours("1.0,-0,1.50,1e21\na,b,c,d\n", { header: "first-row" })
    expect(Object.keys(record!)).toEqual(["0", "1", "1.5", "1e+21"])
  })
})

describe.skipIf(xlsx === undefined)("toRecords vs SheetJS oracle", () => {
  const DIFFERENT: Record<string, readonly [string, Options, unknown, unknown]> = {
    "1. __EMPTY numbering follows first use": [
      "a,,b,,c\n1,,3,4,5\n,2,,,\n",
      { header: "first-row" },
      [{ a: 1, b: 3, __EMPTY: 4, c: 5 }, { __EMPTY_1: 2 }],
      [{ a: 1, b: 3, __EMPTY_1: 4, c: 5 }, { __EMPTY: 2 }]
    ],
    "2. defval doesn't fill blank-header columns": [
      "a,,b\n1\n",
      { header: "first-row", defval: null },
      [{ a: 1, b: null }],
      [{ a: 1, __EMPTY: null, b: null }]
    ],
    "3. repeated explicit names are made unique": ["1,2\n", { header: ["x", "x"] }, [{ x: 1, x_1: 2 }], [{ x: 2 }]],
    "4. an explicit \"\" entry gives __EMPTY": [
      "1,2,3\n",
      { header: ["x", "", "y"] },
      [{ x: 1, __EMPTY: 2, y: 3 }],
      [{ x: 1, "": 2, y: 3 }]
    ],
    "5. numeric header cells are named from the parsed number": [
      "1.0,-0,1.50,1e21\na,b,c,d\n",
      { header: "first-row" },
      [{ "1": "a", "0": "b", "1.5": "c", "1e+21": "d" }],
      [{ "1.0": "a", "-0": "b", "1.50": "c", "1e21": "d" }]
    ],
    "5. numeric header cells with the same number collide": [
      "1,1.0,x\na,b,c\n",
      { header: "first-row" },
      [{ "1": "a", "1_1": "b", x: "c" }],
      [{ "1": "a", "1.0": "b", x: "c" }]
    ]
  }
  for (const [name, [text, options, expectedOurs, expectedTheirs]] of Object.entries(DIFFERENT)) {
    it(`intended difference ${name}`, async () => {
      expect(await ours(text, options)).toEqual(expectedOurs)
      expect(theirs(text, options)).toEqual(expectedTheirs)
    })
  }

  const SAME: Record<string, readonly [string, Options]> = {
    "plain header": ["name,age\nann,31\nbob,\n", { header: "first-row" }],
    "repeated header-row names": ["a,a,b,a\n1,2,3,4\n", { header: "first-row" }],
    "TRUE/FALSE header cells": ["TRUE,FALSE,x\n1,2,3\n", { header: "first-row" }],
    "explicit names, extra columns dropped": ["1,2,3\n4,5\n", { header: ["x", "y"] }],
    "defval with a dense header": ["a,b\n1\n,2\n", { header: "first-row", defval: null }],
    "plain integer header cells": ["1,2,x\na,b,c\n", { header: "first-row" }]
  }
  for (const [name, [text, options]] of Object.entries(SAME)) {
    it(`same records: ${name}`, async () => {
      expect(await ours(text, options)).toEqual(theirs(text, options))
    })
  }

  // A styled blank cell (read as Empty here, absent in SheetJS) under a blank
  // header cell: not data, so it neither takes __EMPTY nor gets defval.
  describe("a styled blank cell under a blank header cell", () => {
    const data = `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>` +
      `<row r="2"><c r="A2"><v>1</v></c><c r="B2" s="1"/></row>` +
      `<row r="3"><c r="A3"><v>2</v></c><c r="D3" t="s"><v>2</v></c></row>`
    // SheetJS finds the parts through [Content_Types].xml.
    const types = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `</Types>`
    const book = () =>
      buildXlsx({
        sheets: [{ name: "S", data }],
        sharedStrings: ["a", "c", "x"],
        xfs: [0, 0],
        extra: [{ name: "[Content_Types].xml", data: enc(types) }]
      })
    const oursXlsx = async (options: Options) => {
      const bytes = await book()
      return Effect.runPromise(Effect.gen(function*() {
        const wb = yield* Xlsx.read(bytes)
        const rows = Array.from(yield* Stream.runCollect(Workbook.rows(wb, 0)))
        expect(rows[1]!.cells.map((c) => c._tag)).toEqual(["Number", "Empty"])
        return Array.from(yield* Stream.runCollect(Spreadsheet.toRecords(Workbook.rows(wb, 0), options)))
      }))
    }
    const theirsXlsx = async (options: Options) => {
      const wb = xlsx!.read(await book(), { type: "buffer" })
      const o: Record<string, unknown> = { raw: true }
      if ("defval" in options) o["defval"] = options.defval
      return xlsx!.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]!], o)
    }
    it("intended difference 1 without defval: the data in column D is the first __EMPTY", async () => {
      expect(await oursXlsx({ header: "first-row" })).toEqual([{ a: 1 }, { a: 2, __EMPTY: "x" }])
      expect(await theirsXlsx({ header: "first-row" })).toEqual([{ a: 1 }, { a: 2, __EMPTY_1: "x" }])
    })
    it("intended difference 2 with defval: only the header-named columns are filled", async () => {
      expect(await oursXlsx({ header: "first-row", defval: null })).toEqual([
        { a: 1, c: null },
        { a: 2, c: null, __EMPTY: "x" }
      ])
      expect(await theirsXlsx({ header: "first-row", defval: null })).toEqual([
        { a: 1, __EMPTY: null, c: null, __EMPTY_1: null },
        { a: 2, __EMPTY: null, c: null, __EMPTY_1: "x" }
      ])
    })
  })
})

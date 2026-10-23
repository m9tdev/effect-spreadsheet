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
import { enc } from "./helpers/run.ts"
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
})

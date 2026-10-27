/**
 * XLSB vs the committed SheetJS expectations (test/corpus/expected, made by
 * scripts/corpus; no SheetJS needed here). For every XLSB fixture: sheet
 * names and visibility, every cell's type and value (dates by serial),
 * hidden rows and columns, merges, the dimension, the cell count and the
 * defined names (name, scope, formula text, hidden). Places listed in
 * test/corpus/sheetjs-quirks.json use their `correct` value. Intended
 * differences are in docs/xlsb-oracle-differences.md.
 */
import { readdirSync, readFileSync } from "node:fs"
import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import type * as Cell from "../src/Cell.ts"
import { isDateCode } from "../src/internal/xlsx/numFmt.ts"
import { columnName } from "../src/internal/xlsx/refs.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as Workbook from "../src/Workbook.ts"

interface OracleCell {
  readonly r: number
  readonly c: number
  readonly t: string
  readonly v?: unknown
  readonly w?: string
  readonly z?: string
  readonly f?: string
}

interface OracleName {
  readonly name: string
  readonly ref: unknown
  readonly sheet?: number
  readonly hidden?: boolean
}

interface Expected {
  readonly status: string
  readonly path: string
  readonly workbook: {
    readonly date1904: boolean
    readonly sheets: ReadonlyArray<{ name: string; visibility: string }>
    /** SheetJS's wb.Workbook.Names, only in files that have some. */
    readonly names?: ReadonlyArray<OracleName>
  }
  readonly sheets: ReadonlyArray<{
    readonly name: string
    readonly hiddenRows: ReadonlyArray<number>
    readonly hiddenCols: ReadonlyArray<number>
    readonly merges: ReadonlyArray<string>
    readonly ref?: string
    readonly cellCount: number
    readonly cells: ReadonlyArray<OracleCell>
  }>
}

interface Quirk {
  readonly file: string
  readonly sheet: string
  readonly field: string
  readonly correct: ReadonlyArray<readonly [number, number, unknown]>
}

const quirks = (JSON.parse(readFileSync("test/corpus/sheetjs-quirks.json", "utf8")) as { quirks: ReadonlyArray<Quirk> }).quirks

/**
 * Defined names where we differ from SheetJS on purpose (docs/xlsb-oracle-differences.md),
 * by fixture and name: SheetJS's value, ours, and why. Each entry must match both readers
 * and name a name SheetJS reports, so the list can't go stale.
 */
interface NameDifference {
  readonly formula?: readonly [sheetjs: unknown, ours: string]
  readonly hidden?: readonly [sheetjs: boolean | undefined, ours: boolean]
  readonly why: string
}
const nameDifferences: Record<string, Record<string, NameDifference>> = {
  "generated/names_formulas.xlsb": {
    Span: { formula: ["'Jan:Mar'!$A$1:$A$10", "Jan:Mar!$A$1:$A$10"], why: "a sheet span is quoted only when a sheet name needs it, as Excel writes it" },
    Accents: { formula: ["'Données'!$A$1", "Données!$A$1"], why: "letters beyond ASCII don't need quotes in Excel" },
    Rate: { formula: ["SH33TJSSRC355!$A$1", "[1]Rates!$A$1"], why: "an external workbook keeps its [n] and its sheet name, from its external link part (SheetJS writes a placeholder)" },
    Lost: { formula: ["#REF!", "Data!#REF!"], why: "PtgRefErr3d keeps its sheet, as Excel writes it" },
    LostArea: { formula: ["#REF!", "Data!#REF!"], why: "PtgAreaErr3d keeps its sheet, as Excel writes it" },
    Calc: { formula: [42, ""], why: "formulas other than a single reference or area aren't decompiled" },
    Secret: { hidden: [undefined, true], why: "the hidden flag comes from the BrtName's fHidden; SheetJS doesn't report it" },
    "_xlnm._FilterDatabase": { hidden: [undefined, true], why: "the hidden flag comes from the BrtName's fHidden; SheetJS doesn't report it" }
  }
}

const errorText = new Map([[0x00, "#NULL!"], [0x07, "#DIV/0!"], [0x0F, "#VALUE!"], [0x17, "#REF!"], [0x1D, "#NAME?"], [0x24, "#NUM!"], [0x2A, "#N/A"], [0x2B, "#GETTING_DATA"]])

const DAY_MS = 86_400_000
/** Our Date back to a serial, the inverse of serialToDate. */
const toSerial = (date: Date, date1904: boolean): number => {
  const ms = date.getTime()
  if (date1904) return (ms - Date.UTC(1904, 0, 1)) / DAY_MS
  const serial = (ms - Date.UTC(1899, 11, 30)) / DAY_MS
  return serial < 61 ? (ms - Date.UTC(1899, 11, 31)) / DAY_MS : serial
}

/** Serials that are no real date, so they stay numbers for us (docs/xlsb-oracle-differences.md). */
const notADate = (serial: number, date1904: boolean): boolean =>
  serial < 0 || serial >= (date1904 ? 2_958_466 - 1462 : 2_958_466) || (!date1904 && serial >= 60 && serial < 61)

const a1 = (r: Workbook.Range): string => {
  const first = `${columnName(r.firstColumn)}${r.firstRow + 1}`
  const last = `${columnName(r.lastColumn)}${r.lastRow + 1}`
  return first === last ? first : `${first}:${last}`
}

const fixtures = [
  ...readdirSync("test/corpus/expected/generated").filter((n) => n.endsWith(".xlsb.json")).map((n) => ["generated", n] as const),
  ...readdirSync("test/corpus/expected/hidden-rows").filter((n) => n.endsWith(".xlsb.json")).map((n) => ["hidden-rows", n] as const)
]

const readAll = (bytes: Uint8Array) =>
  Effect.runPromise(Effect.gen(function*() {
    const wb = yield* Spreadsheet.read(bytes)
    const sheets: Array<ReadonlyArray<Workbook.Row>> = []
    const details: Array<Workbook.SheetDetails> = []
    for (const sheet of wb.sheets) {
      sheets.push(Array.from(yield* Stream.runCollect(Workbook.rows(wb, sheet.index))))
      details.push(yield* Workbook.sheetDetails(wb, sheet.index))
    }
    return { wb, sheets, details }
  }))

describe("xlsb vs the committed SheetJS expectations", () => {
  it("covers every committed xlsb fixture", () => {
    expect(fixtures.length).toBeGreaterThanOrEqual(6)
    const files = fixtures.map(([source, name]) => `${source}/${name.replace(/\.json$/, "")}`)
    expect(Object.keys(nameDifferences).filter((f) => !files.includes(f))).toEqual([])
  })

  for (const [source, name] of fixtures) {
    it(`${source}/${name.replace(/\.json$/, "")}`, async () => {
      const expected = JSON.parse(readFileSync(`test/corpus/expected/${source}/${name}`, "utf8")) as Expected
      expect(expected.status).toBe("ok")
      const file = `${source}/${name.replace(/\.json$/, "")}`
      const { wb, sheets, details } = await readAll(new Uint8Array(readFileSync(`test/corpus/${file}`)))
      expect(wb.format).toBe("xlsb")
      expect(wb.sheets.map((s) => ({ name: s.name, visibility: s.visibility }))).toEqual(expected.workbook.sheets)
      // defined names: name, scope, formula (SheetJS's without a leading "=") and hidden, in file order
      const differences = nameDifferences[file] ?? {}
      const oracleNames = (expected.workbook.names ?? []).map((n) => {
        const ref = typeof n.ref === "string" ? n.ref.replace(/^=/, "") : n.ref
        const d = differences[n.name]
        if (d?.formula !== undefined) expect(ref, `${n.name}: SheetJS formula`).toEqual(d.formula[0])
        if (d?.hidden !== undefined) expect(n.hidden, `${n.name}: SheetJS hidden`).toBe(d.hidden[0])
        return {
          name: n.name,
          formula: d?.formula !== undefined ? d.formula[1] : ref,
          ...(n.sheet === undefined ? {} : { scope: n.sheet }),
          hidden: d?.hidden !== undefined ? d.hidden[1] : n.hidden ?? false
        }
      })
      expect(wb.definedNames).toEqual(oracleNames)
      // every listed difference is about a name SheetJS reports, so the list can't go stale
      const reported = new Set((expected.workbook.names ?? []).map((n) => n.name))
      expect(Object.keys(differences).filter((n) => !reported.has(n)), "unused name differences").toEqual([])
      const date1904 = expected.workbook.date1904
      expected.sheets.forEach((oracle, k) => {
        const rows = sheets[k]!
        const ours = new Map<string, Cell.Cell>()
        for (const row of rows) for (const cell of row.cells) ours.set(`${row.index}:${cell.column}`, cell)
        const corrections = new Map<string, unknown>()
        for (const q of quirks) {
          if (q.file === file && q.sheet === oracle.name && q.field === "cells") {
            for (const [r, c, v] of q.correct) corrections.set(`${r}:${c}`, v)
          }
        }
        for (const cell of oracle.cells) {
          const key = `${cell.r}:${cell.c}`
          const mine = ours.get(key)
          const where = `${oracle.name}!${key}`
          expect(mine, where).toBeDefined()
          const value = corrections.has(key) ? corrections.get(key) : cell.v
          switch (cell.t) {
            case "z":
              // a styled blank cell (BrtCellBlank): SheetJS gives a stub for XLSB even without sheetStubs
              expect(mine!._tag, where).toBe("Empty")
              break
            case "s":
              expect([mine!._tag, (mine as Cell.TextCell).value], where).toEqual(["Text", value])
              break
            case "b":
              expect([mine!._tag, (mine as Cell.BooleanCell).value], where).toEqual(["Boolean", value])
              break
            case "e":
              // SheetJS gives no `w` for a formula's error result: name its code (MS-XLS BErr)
              expect([mine!._tag, (mine as Cell.ErrorCell).value], where).toEqual(["Error", cell.w ?? errorText.get(value as number)])
              break
            case "n":
              if (mine!._tag === "Date") {
                // a date for us: the oracle has the serial and a date format
                expect(cell.z !== undefined && isDateCode(cell.z), `${where} z=${cell.z}`).toBe(true)
                expect(Math.abs(toSerial(mine!.value, date1904) - (value as number)), where).toBeLessThan(1e-8)
              } else {
                expect([mine!._tag, (mine as Cell.NumberCell).value], where).toEqual(["Number", value])
                if (cell.z !== undefined && isDateCode(cell.z)) expect(notADate(value as number, date1904), `${where} ${cell.z}`).toBe(true)
              }
              break
            default:
              throw new Error(`unexpected oracle type ${cell.t} at ${where}`)
          }
          // formulas aren't decompiled: the cached result only (docs/xlsb-oracle-differences.md)
          expect(mine!.formula, where).toBeUndefined()
        }
        // nothing beyond the oracle's cells, except Empty ones (styled blanks, stubs for SheetJS)
        const valued = rows.flatMap((r) => r.cells).filter((c) => c._tag !== "Empty")
        expect(valued.length, oracle.name).toBe(oracle.cells.filter((c) => c.t !== "z").length)
        expect(rows.filter((r) => r.hidden).map((r) => r.index), oracle.name).toEqual(oracle.hiddenRows)
        // layout: hidden columns, merges, and the declared dimension where the file has one
        const layout = details[k]!
        expect(layout.hiddenColumns.flatMap((s) => Array.from({ length: s.last - s.first + 1 }, (_, i) => s.first + i)), oracle.name).toEqual(oracle.hiddenCols)
        expect(layout.merges.map(a1), oracle.name).toEqual(oracle.merges)
        if (layout.dimension !== undefined) expect(a1(layout.dimension), oracle.name).toBe(oracle.ref)
      })
    })
  }
})

describe("hidden-row fixtures (the reason this project exists)", () => {
  const manifest = JSON.parse(readFileSync("test/corpus/hidden-rows/manifest.json", "utf8")) as {
    files: ReadonlyArray<{
      path: string
      format: string
      expected: { sheets: ReadonlyArray<{ name: string; rows: ReadonlyArray<{ index: number; hidden: boolean; values: ReadonlyArray<unknown> }> }> }
    }>
  }
  for (const file of manifest.files.filter((f) => f.format === "xlsb")) {
    it(`${file.path}: every row, hidden ones included, whatever the file is called`, async () => {
      const { wb, sheets } = await readAll(new Uint8Array(readFileSync(`test/corpus/hidden-rows/${file.path}`)))
      expect(wb.format).toBe("xlsb")
      file.expected.sheets.forEach((sheet, k) => {
        expect(wb.sheets[k]!.name).toBe(sheet.name)
        const rows = sheets[k]!.map((r) => ({ index: r.index, hidden: r.hidden, values: r.cells.map((c) => (c._tag === "Empty" ? null : c.value)) }))
        expect(rows).toEqual(sheet.rows)
      })
    })
  }
})

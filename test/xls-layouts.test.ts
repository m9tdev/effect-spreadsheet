/**
 * XLS record layouts: every record the reader decodes, cut at every length.
 * Below its declared minimum a cell, ROW, BOUNDSHEET, SST, COLINFO or
 * MERGECELLS record is MalformedFile; a cut CODEPAGE, DATEMODE, FORMAT, XF,
 * DIMENSIONS, NAME or EXTERNSHEET is tolerated, each in its own way.
 * Synthetic records only.
 */
import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import * as Cell from "../src/Cell.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { failureOf } from "./helpers/run.ts"
import * as X from "./helpers/xls.ts"

const readExit = (book: X.TestBook, limits?: LimitOverrides) =>
  Effect.runPromise(Effect.exit(Effect.gen(function*() {
    const wb = yield* Xls.read(X.xls(book), { limits })
    const sheets: Array<Array<Workbook.Row>> = []
    const details: Array<Workbook.SheetDetails> = []
    for (const sheet of wb.sheets) {
      sheets.push(Array.from(yield* Stream.runCollect(Workbook.rows(wb, sheet.index))))
      details.push(yield* Workbook.sheetDetails(wb, sheet.index))
    }
    return { wb, sheets, details }
  })))
const readOk = async (book: X.TestBook) => {
  const exit = await readExit(book)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value
}
const errorOf = async (book: X.TestBook) => failureOf(await readExit(book))
const one = (records: ReadonlyArray<Uint8Array>, globals: ReadonlyArray<Uint8Array> = [], biff: 5 | 8 = 8): X.TestBook => ({ biff, globals, sheets: [{ name: "S", records }] })
/** A record's data, without its 4-byte header. */
const dataOf = (record: Uint8Array) => Array.from(record.subarray(4))
/** `record` with its data cut to `length` bytes. */
const cutTo = (record: Uint8Array, length: number) => X.record(record[0]! | (record[1]! << 8), dataOf(record).slice(0, length))
/** The values of the first sheet's cells. */
const values = async (book: X.TestBook) => (await readOk(book)).sheets[0]!.flatMap((r) => r.cells.map((c) => Cell.value(c)))

describe("xls record layouts: sheet records", () => {
  it("every cell record and ROW cut below its declared minimum, or inside its string, is a truncated record", async () => {
    // a record, the length below which it is truncated (its declared minimum, or its whole length when a string ends it)
    const kinds: ReadonlyArray<readonly [Uint8Array, number]> = [
      [X.row(0), 14],
      [X.labelSst(0, 0, 0, 0), 10],
      [X.label(0, 0, 0, "label"), X.label(0, 0, 0, "label").length - 4],
      [X.record(0x00D6, [...X.u16(0), ...X.u16(0), ...X.u16(0), ...X.xlString("rich"), ...X.u16(0)]), 6 + X.xlString("rich").length],
      [X.number(0, 0, 0, 2.5), 14],
      [X.rk(0, 0, 0, X.rkInt(7)), 10],
      [X.mulrk(0, 0, [[0, X.rkInt(1)], [0, X.rkInt(2)]]), 6],
      [X.mulBlank(0, 0, [0, 0]), 6],
      [X.boolErr(0, 0, 0, 1, false), 8],
      [X.blank(0, 0, 0), 6],
      [X.formula(0, 0, 0, 1.5), 14],
      [X.record(0x0406, [...X.u16(0), ...X.u16(0), ...X.u16(0), ...X.f64(1.5), ...X.u16(0), ...X.u32(0)]), 14]
    ]
    // progress: one kind per iteration
    for (const [record, min] of kinds) {
      const type = record[0]! | (record[1]! << 8)
      expect(await errorOf(one([X.row(0), record], [X.sst(["a"])])), `whole record 0x${type.toString(16)}`).toBeUndefined()
      // progress: k up to the record's length
      for (let k = 0; k < dataOf(record).length; k++) {
        const error = await errorOf(one([X.row(0), cutTo(record, k)], [X.sst(["a"])]))
        if (k < min) {
          expect(error, `record 0x${type.toString(16)} cut to ${k} bytes`)
            .toMatchObject({ _tag: "MalformedFile", reason: `truncated record 0x${type.toString(16)}` })
        } else {
          expect(error?._tag === "MalformedFile" && error.reason, `record 0x${type.toString(16)} cut to ${k} bytes`)
            .not.toBe(`truncated record 0x${type.toString(16)}`)
        }
      }
    }
  })

  it("a STRING cut anywhere, in BIFF8 or BIFF5, is a truncated record", async () => {
    const biff8 = X.string("text")
    const biff5 = X.record(0x0207, X.string5([0x61, 0x62, 0x63], 2))
    for (const [record, biff] of [[biff8, 8], [biff5, 5]] as const) {
      expect(await values(one([X.formula(0, 0, 0, { kind: 0 }), record], [], biff))).toEqual([biff === 8 ? "text" : "abc"])
      // progress: k up to the record's length
      for (let k = 0; k < dataOf(record).length; k++) {
        expect(await errorOf(one([X.formula(0, 0, 0, { kind: 0 }), cutTo(record, k)], [], biff)), `BIFF${biff} STRING cut to ${k} bytes`)
          .toMatchObject({ _tag: "MalformedFile", reason: "truncated record 0x207" })
      }
    }
  })

  it("a cut COLINFO or MERGECELLS is MalformedFile; a cut DIMENSIONS gives no dimension", async () => {
    const colInfo = X.colInfo(1, 2, true)
    expect((await readOk(one([colInfo]))).details[0]!.hiddenColumns).toEqual([{ first: 1, last: 2 }])
    // progress: k up to 10 (colFirst, colLast, coldx, ixfe, flags)
    for (let k = 0; k < 10; k++) {
      expect(await errorOf(one([cutTo(colInfo, k)])), `COLINFO cut to ${k} bytes`).toMatchObject({ _tag: "MalformedFile", reason: "truncated COLINFO" })
    }
    expect((await readOk(one([cutTo(colInfo, 10)]))).details[0]!.hiddenColumns).toEqual([{ first: 1, last: 2 }])
    const merges = X.mergeCells([[0, 1, 0, 1], [2, 3, 2, 3]])
    expect((await readOk(one([merges]))).details[0]!.merges).toHaveLength(2)
    // progress: k up to the record's length
    for (let k = 0; k < dataOf(merges).length; k++) {
      expect(await errorOf(one([cutTo(merges, k)])), `MERGECELLS cut to ${k} bytes`).toMatchObject({ _tag: "MalformedFile", reason: "truncated MERGECELLS" })
    }
    for (const [biff, min] of [[8, 12], [5, 8]] as const) {
      const dimensions = X.dimensions(0, 3, 0, 2, biff)
      const dimensionOf = async (record: Uint8Array) => (await readOk(one([record], [], biff))).details[0]!.dimension
      expect(await dimensionOf(dimensions)).toEqual({ firstRow: 0, firstColumn: 0, lastRow: 2, lastColumn: 1 })
      expect(await dimensionOf(cutTo(dimensions, min))).toEqual({ firstRow: 0, firstColumn: 0, lastRow: 2, lastColumn: 1 })
      // progress: k up to the declared minimum
      for (let k = 0; k < min; k++) expect(await dimensionOf(cutTo(dimensions, k)), `BIFF${biff} DIMENSIONS cut to ${k} bytes`).toBeUndefined()
    }
  })
})

describe("xls record layouts: globals records", () => {
  const dateCell = [X.row(0), X.number(0, 0, 1, 1)]

  it("a cut XF still takes its index, with the General format", async () => {
    const xf = X.xf(14)
    expect(await values(one(dateCell, [X.xf(0), xf]))).toEqual([new Date(Date.UTC(1900, 0, 1))])
    expect(await values(one(dateCell, [X.xf(0), cutTo(xf, 4)]))).toEqual([new Date(Date.UTC(1900, 0, 1))])
    // progress: k up to 4 (ifnt, ifmt)
    for (let k = 0; k < 4; k++) expect(await values(one(dateCell, [X.xf(0), cutTo(xf, k)])), `XF cut to ${k} bytes`).toEqual([1])
  })

  it("a cut FORMAT is ignored, in BIFF8 and BIFF5", async () => {
    for (const biff of [8, 5] as const) {
      const format = X.format(164, "yyyy-mm-dd", biff)
      expect(await values(one(dateCell, [X.xf(0), X.xf(164), format], biff))).toEqual([new Date(Date.UTC(1900, 0, 1))])
      // progress: k up to the record's length
      for (let k = 0; k < dataOf(format).length; k++) {
        expect(await values(one(dateCell, [X.xf(0), X.xf(164), cutTo(format, k)], biff)), `BIFF${biff} FORMAT cut to ${k} bytes`).toEqual([1])
      }
    }
  })

  it("a cut DATEMODE means the 1900 date system; a cut CODEPAGE means windows-1252", async () => {
    const datemode = X.datemode1904()
    expect(await values(one(dateCell, [datemode, X.xf(0), X.xf(14)]))).toEqual([new Date(Date.UTC(1904, 0, 2))])
    // progress: k up to 2
    for (let k = 0; k < 2; k++) {
      expect(await values(one(dateCell, [cutTo(datemode, k), X.xf(0), X.xf(14)])), `DATEMODE cut to ${k} bytes`).toEqual([new Date(Date.UTC(1900, 0, 1))])
    }
    const codepage = X.codepage(1251)
    const nameWith = async (record: Uint8Array) =>
      (await readOk({ biff: 5, globals: [record], sheets: [{ name: "", name5: [0xC0], records: [] }] })).wb.sheets[0]!.name
    expect(await nameWith(codepage)).toBe("\u0410")
    // progress: k up to 2
    for (let k = 0; k < 2; k++) expect(await nameWith(cutTo(codepage, k)), `CODEPAGE cut to ${k} bytes`).toBe("\u00C0")
  })

  it("a cut BOUNDSHEET or SST header is MalformedFile", async () => {
    const boundsheet = X.record(0x0085, [...X.u32(0), 0, 0, ...X.shortString("Lost")])
    // progress: k up to the record's length
    for (let k = 0; k < dataOf(boundsheet).length; k++) {
      expect(await errorOf(one([], [cutTo(boundsheet, k)])), `BOUNDSHEET cut to ${k} bytes`)
        .toMatchObject({ _tag: "MalformedFile", reason: "truncated BOUNDSHEET" })
    }
    const sst = X.sst(["a"])
    // progress: k up to 8 (cstTotal, cstUnique)
    for (let k = 0; k < 8; k++) {
      expect(await errorOf(one([], [cutTo(sst, k)])), `SST cut to ${k} bytes`).toMatchObject({ _tag: "MalformedFile", reason: "truncated SST" })
    }
    // the strings cut: those before the cut are kept
    expect(await values(one([X.labelSst(0, 0, 0, 0)], [cutTo(sst, 8 + X.xlString("a").length)]))).toEqual(["a"])
    expect(await errorOf(one([X.labelSst(0, 0, 0, 0)], [cutTo(sst, 8 + X.xlString("a").length - 1)])))
      .toMatchObject({ _tag: "MalformedFile", reason: "shared string 0 doesn't exist" })
  })

  it("a NAME cut in its header, name or formula is skipped; an EXTERNSHEET keeps its whole XTIs", async () => {
    const self = X.record(0x01AE, [...X.u16(1), 0x01, 0x04])
    const names = async (globals: ReadonlyArray<Uint8Array>) =>
      Object.fromEntries((await readOk(one([], [self, ...globals]))).wb.definedNames.map((d) => [d.name, d.formula]))
    const name = X.name("N", X.ref3d(0, 0, 0))
    const extern = X.externSheet([0, 0])
    expect(await names([extern, name])).toEqual({ N: "S!$A$1" })
    // progress: k up to the record's length
    for (let k = 0; k < dataOf(name).length; k++) expect(await names([extern, cutTo(name, k)]), `NAME cut to ${k} bytes`).toEqual({})
    // an empty name and formula: only its 15-byte header is read
    const empty = X.name("", [])
    expect(await names([X.name("A", []), empty, X.name("B", [])])).toEqual({ A: "", "": "", B: "" })
    // progress: k up to 15
    for (let k = 0; k < 15; k++) expect(await names([X.name("A", []), cutTo(empty, k), X.name("B", [])]), `empty NAME cut to ${k} bytes`).toEqual({ A: "", B: "" })
    // progress: k up to the record's length: XTI 1 is read once its 6 bytes are there
    for (let k = 0; k < dataOf(extern).length; k++) {
      const second = X.name("Second", X.ref3d(1, 0, 0))
      expect(await names([cutTo(extern, k), second]), `EXTERNSHEET cut to ${k} bytes`).toEqual({ Second: "" })
    }
  })
})

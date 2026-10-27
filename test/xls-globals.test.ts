/** XLS workbook globals: defined-name text (SUPBOOK/EXTERNSHEET), BOF versions, sheet visibility, SST, text decoding, containers. */
import { readFileSync } from "node:fs"
import { Effect, Exit, Stream } from "effect"
import { afterEach, describe, expect, it, vi } from "vitest"
import * as Cell from "../src/Cell.ts"
import { readString, type Text, textDecoder } from "../src/internal/xls/fields.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import * as Xls from "../src/xls/index.ts"
import { writeCfb } from "./helpers/cfb.ts"
import { failureOf } from "./helpers/run.ts"
import * as X from "./helpers/xls.ts"

const run = (bytes: Uint8Array, limits?: LimitOverrides) => Effect.runPromise(Effect.exit(Xls.read(bytes, { limits })))
const read = async (book: X.TestBook, limits?: LimitOverrides) => {
  const exit = await run(X.xls(book), limits)
  if (Exit.isFailure(exit)) throw new Error(`read failed: ${JSON.stringify(failureOf(exit))}`)
  return exit.value
}
const errorOf = async (book: X.TestBook, limits?: LimitOverrides) => failureOf(await run(X.xls(book), limits))
const texts = async (wb: Workbook.Workbook, sheet = 0) =>
  Array.from(await Effect.runPromise(Stream.runCollect(Workbook.rows(wb, sheet)))).flatMap((r) => r.cells.map((c) => Cell.value(c)))

// SUPBOOK records: this workbook (ctab, 0x0401), an add-in (0x3A01), an external workbook (ctab, path, sheet names)
const SUPBOOK = 0x01AE
const selfBook = (sheets: number) => X.record(SUPBOOK, [...X.u16(sheets), 0x01, 0x04])
const addinBook = () => X.record(SUPBOOK, [...X.u16(1), 0x01, 0x3A])
const externalBook = (path: string, sheets: ReadonlyArray<string>) => X.record(SUPBOOK, [...X.u16(sheets.length), ...X.xlString(path), ...sheets.flatMap(X.xlString)])
/** EXTERNSHEET with explicit XTIs: [supbook, itabFirst, itabLast]. */
const xtis = (entries: ReadonlyArray<readonly [number, number, number]>) =>
  X.record(0x0017, [...X.u16(entries.length), ...entries.flatMap(([b, f, l]) => [...X.u16(b), ...X.u16(f), ...X.u16(l)])])
const CONTINUE = 0x003C

/** PtgRef3d / PtgArea3d in a given class (0x3A/0x3B reference, 0x5A/0x5B value, 0x7A/0x7B array). */
const ref3d = (ptg: number, ixti: number, r: number, c: number) => [ptg, ...X.u16(ixti), ...X.u16(r), ...X.u16(c)]
const area3d = (ptg: number, ixti: number, r1: number, r2: number, c1: number, c2: number) => [ptg, ...X.u16(ixti), ...X.u16(r1), ...X.u16(r2), ...X.u16(c1), ...X.u16(c2)]
const COL_REL = 0x4000
const ROW_REL = 0x8000

const formulas = (wb: Workbook.Workbook) => Object.fromEntries(wb.definedNames.map((d) => [d.name, d.formula]))
const sheets = (...names: ReadonlyArray<string>) => names.map((name) => ({ name, records: [] }))

describe("xls defined names: which workbook a reference points at", () => {
  it("only XTIs of the self SUPBOOK resolve to local sheets; external ones name the link, deleted sheets are #REF!", async () => {
    const wb = await read({
      globals: [
        selfBook(2),
        externalBook("\u0001other.xls", ["Ext One", "Data"]),
        xtis([[0, 1, 1], [1, 0, 0], [1, 1, 1], [0, 0xFFFF, 0xFFFF], [1, 0xFFFE, 0xFFFE], [0, 5, 5], [1, 0, 1]]),
        X.name("Local", ref3d(0x3A, 0, 0, 0)),
        X.name("ExtQuoted", ref3d(0x3A, 1, 1, 1)),
        X.name("ExtPlain", area3d(0x3B, 2, 0, 1, 0, 1)),
        X.name("Deleted", ref3d(0x3A, 3, 2, 0)),
        X.name("ExtBookLevel", ref3d(0x3A, 4, 0, 0)),
        X.name("OutOfRange", ref3d(0x3A, 5, 0, 0)),
        X.name("ExtSpan", ref3d(0x3A, 6, 0, 0))
      ],
      sheets: sheets("Data", "Mine")
    })
    expect(formulas(wb)).toEqual({
      Local: "Mine!$A$1",
      ExtQuoted: "'[1]Ext One'!$B$2",
      ExtPlain: "[1]Data!$A$1:$B$2",
      Deleted: "#REF!$A$3",
      ExtBookLevel: "#REF!$A$1",
      OutOfRange: "#REF!$A$1",
      ExtSpan: "'[1]Ext One:Data'!$A$1"
    })
  })

  it("the self SUPBOOK need not come first; external links are numbered in record order, add-ins skipped", async () => {
    const wb = await read({
      globals: [
        externalBook("\u0001a.xls", ["A"]),
        addinBook(),
        externalBook("\u0001b.xls", ["B"]),
        selfBook(1),
        xtis([[3, 0, 0], [0, 0, 0], [2, 0, 0], [1, 0, 0]]),
        X.name("Here", ref3d(0x3A, 0, 0, 0)),
        X.name("InA", ref3d(0x3A, 1, 0, 0)),
        X.name("InB", ref3d(0x3A, 2, 0, 0)),
        X.name("AddIn", ref3d(0x3A, 3, 0, 0))
      ],
      sheets: sheets("Local")
    })
    expect(formulas(wb)).toEqual({ Here: "Local!$A$1", InA: "[1]A!$A$1", InB: "[2]B!$A$1", AddIn: "#REF!$A$1" })
  })

  it("a 3-D span of local sheets (itabFirst ≠ itabLast) names both ends; a reversed span is #REF!", async () => {
    const wb = await read({
      globals: [
        selfBook(3),
        xtis([[0, 0, 2], [0, 1, 2], [0, 2, 0]]),
        X.name("Span", area3d(0x3B, 0, 0, 0, 0, 0)),
        X.name("SpanQuoted", ref3d(0x3A, 1, 0, 0)),
        X.name("Reversed", ref3d(0x3A, 2, 0, 0))
      ],
      sheets: sheets("Jan", "Feb", "Mar 2")
    })
    expect(formulas(wb)).toEqual({ Span: "'Jan:Mar 2'!$A$1:$A$1", SpanQuoted: "'Feb:Mar 2'!$A$1", Reversed: "#REF!$A$1" })
  })

  it("an external path whose length reads like the self or add-in marker (0x0401, 0x3A01) stays external", async () => {
    // 1025 characters: one record
    const near = externalBook(`\u0001${"p".repeat(1024)}`, ["Far"])
    // 0x3A01 = 14 849 characters: more than one record holds, the rest in a CONTINUE (its own option byte first)
    const path = [...X.u16(0x3A01), 0x00, ...new Array(0x3A01).fill(0x71)]
    const addinLike = [...X.u16(1), ...path, ...X.xlString("Away")]
    const cut = 2 + 3 + 8000
    const wb = await read({
      globals: [
        selfBook(1),
        near,
        X.record(SUPBOOK, addinLike.slice(0, cut)),
        X.record(CONTINUE, [0x00, ...addinLike.slice(cut)]),
        xtis([[1, 0, 0], [2, 0, 0]]),
        X.name("Near", ref3d(0x3A, 0, 0, 0)),
        X.name("AddinLike", ref3d(0x3A, 1, 0, 0))
      ],
      sheets: sheets("S")
    })
    expect(formulas(wb)).toEqual({ Near: "[1]Far!$A$1", AddinLike: "[2]Away!$A$1" })
  })

  it("external sheet names continued in a CONTINUE record are read", async () => {
    const all = Array.from(externalBook("\u0001x.xls", ["First", "Second"]))
    const data = all.slice(4)
    const cut = 2 + X.xlString("\u0001x.xls").length + X.xlString("First").length
    const wb = await read({
      globals: [
        selfBook(1),
        X.record(SUPBOOK, data.slice(0, cut)),
        X.record(CONTINUE, data.slice(cut)),
        xtis([[1, 1, 1]]),
        X.name("Far", ref3d(0x3A, 0, 0, 0))
      ],
      sheets: sheets("S")
    })
    expect(formulas(wb)).toEqual({ Far: "[1]Second!$A$1" })
  })

  it("with no SUPBOOK at all, SUPBOOK 0 is this workbook; a missing XTI gives an empty formula", async () => {
    const wb = await read({
      globals: [xtis([[0, 0, 0], [1, 0, 0]]), X.name("A", ref3d(0x3A, 0, 0, 0)), X.name("B", ref3d(0x3A, 1, 0, 0)), X.name("C", ref3d(0x3A, 9, 0, 0))],
      sheets: sheets("S")
    })
    expect(formulas(wb)).toEqual({ A: "S!$A$1", B: "#REF!$A$1", C: "" })
  })

  it("SUPBOOK strings and sizes are limited", async () => {
    const long = { globals: [selfBook(1), externalBook("\u0001p.xls", ["x".repeat(40)])], sheets: sheets("S") }
    expect(await errorOf(long, { maxStringLength: 20 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
    const names = Array.from({ length: 100 }, (_, k) => `Sheet${k}`)
    const big = { globals: [X.record(SUPBOOK, [...X.u16(100), ...X.xlString("\u0001p.xls")]), X.record(CONTINUE, names.flatMap(X.xlString))], sheets: sheets("S") }
    expect(await errorOf(big, { maxDecompressedBytes: 200 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes" })
    // a SUPBOOK declaring more sheets than it holds keeps those present
    const short = X.record(SUPBOOK, [...X.u16(1000), ...X.xlString("\u0001p.xls"), ...X.xlString("Only")])
    const wb = await read({ globals: [short, xtis([[0, 0, 0], [0, 1, 1]]), X.name("A", ref3d(0x3A, 0, 0, 0)), X.name("B", ref3d(0x3A, 1, 0, 0))], sheets: sheets("S") })
    expect(formulas(wb)).toEqual({ A: "[1]Only!$A$1", B: "#REF!$A$1" })
  })
})

describe("xls defined names: reference text", () => {
  it("references to deleted cells (PtgRefErr3d, PtgAreaErr3d, every class) are Sheet!#REF!, or #REF! with no sheet", async () => {
    const err = (ptg: number, ixti: number) => [ptg, ...X.u16(ixti), ...new Array(ptg % 16 === 0x0C ? 4 : 8).fill(0)]
    const wb = await read({
      globals: [
        selfBook(2),
        externalBook("\u0001x.xls", ["Ext"]),
        xtis([[0, 0, 0], [0, 1, 1], [0, 0xFFFF, 0xFFFF], [1, 0, 0]]),
        ...[0x3C, 0x5C, 0x7C, 0x3D, 0x5D, 0x7D].map((ptg) => X.name(`P${ptg.toString(16)}`, err(ptg, 0))),
        X.name("Quoted", err(0x3C, 1)),
        X.name("NoSheet", err(0x3D, 2)),
        X.name("External", err(0x3C, 3)),
        X.name("NoXti", err(0x3C, 9)),
        X.name("Trailing", [...err(0x3C, 0), 0x03])
      ],
      sheets: sheets("Data", "My Sheet")
    })
    expect(formulas(wb)).toEqual({
      P3c: "Data!#REF!", P5c: "Data!#REF!", P7c: "Data!#REF!", P3d: "Data!#REF!", P5d: "Data!#REF!", P7d: "Data!#REF!",
      Quoted: "'My Sheet'!#REF!",
      NoSheet: "#REF!",
      External: "[1]Ext!#REF!",
      NoXti: "",
      Trailing: ""
    })
  })

  it("relative flags, every token class, whole rows and columns", async () => {
    const wb = await read({
      globals: [
        selfBook(2),
        xtis([[0, 0, 0], [0, 1, 1]]),
        X.name("Abs", ref3d(0x3A, 1, 4, 1)),
        X.name("RelCol", ref3d(0x5A, 1, 4, 1 | COL_REL)),
        X.name("RelRow", ref3d(0x7A, 1, 4, 1 | ROW_REL)),
        X.name("RelBoth", ref3d(0x3A, 1, 4, 1 | COL_REL | ROW_REL)),
        X.name("Area", area3d(0x3B, 0, 0, 9, 2 | COL_REL, 27 | ROW_REL)),
        X.name("AreaV", area3d(0x5B, 0, 0, 1, 0, 1)),
        X.name("AreaA", area3d(0x7B, 0, 0, 1, 0, 1)),
        X.name("Cols", area3d(0x3B, 0, 0, 0xFFFF, 0, 2)),
        X.name("ColsRel", area3d(0x3B, 0, 0, 0xFFFF, 1 | COL_REL | ROW_REL, 1 | COL_REL)),
        X.name("Rows", area3d(0x3B, 1, 0, 1, 0, 0xFF)),
        X.name("RowsRel", area3d(0x3B, 1, 4, 4, ROW_REL, 0xFF | ROW_REL)),
        X.name("All", area3d(0x3B, 0, 0, 0xFFFF, 0, 0xFF)),
        X.name("NotWhole", area3d(0x3B, 0, 1, 0xFFFF, 0, 0xFE)),
        X.name("LastCol", ref3d(0x3A, 0, 65535, 255)),
        X.name("Trailing", [...ref3d(0x3A, 0, 0, 0), 0x03])
      ],
      sheets: sheets("Data", "My Sheet")
    })
    expect(formulas(wb)).toEqual({
      Abs: "'My Sheet'!$B$5",
      RelCol: "'My Sheet'!B$5",
      RelRow: "'My Sheet'!$B5",
      RelBoth: "'My Sheet'!B5",
      Area: "Data!C$1:$AB10",
      AreaV: "Data!$A$1:$B$2",
      AreaA: "Data!$A$1:$B$2",
      Cols: "Data!$A:$C",
      ColsRel: "Data!B:B",
      Rows: "'My Sheet'!$1:$2",
      RowsRel: "'My Sheet'!5:5",
      All: "Data!$1:$65536",
      NotWhole: "Data!$A$2:$IU$65536",
      LastCol: "Data!$IV$65536",
      Trailing: ""
    })
  })

  it("sheet names are quoted as Excel quotes them, with '' for '", async () => {
    const names = ["Bob's", "Sheet1", "_x.y", "A1", "XFD9", "XFE9", "R", "C3", "R1C1", "RC", "Rate", "1st", "Données", "日本語", "a-b", "x y", ".a"]
    const wb = await read({
      globals: [selfBook(names.length), xtis(names.map((_, k) => [0, k, k] as const)), ...names.map((_, k) => X.name(`N${k}`, ref3d(0x3A, k, 0, 0)))],
      sheets: sheets(...names)
    })
    expect(wb.definedNames.map((d) => d.formula.replace(/!\$A\$1$/, ""))).toEqual([
      "'Bob''s'", "Sheet1", "_x.y", "'A1'", "'XFD9'", "XFE9", "'R'", "'C3'", "'R1C1'", "'RC'", "Rate", "'1st'", "Données", "日本語", "'a-b'", "'x y'", "'.a'"
    ])
  })

  it("an XTI further down the EXTERNSHEET list picks its own sheet; a scope beyond the sheets is dropped", async () => {
    const wb = await read({
      globals: [
        selfBook(3),
        xtis([[0, 2, 2], [0, 0, 0], [0, 1, 1]]),
        X.name("Second", ref3d(0x3A, 2, 0, 0), { itab: 2 }),
        X.name("First", ref3d(0x3A, 1, 0, 0), { itab: 3 }),
        X.name("Lost", ref3d(0x3A, 0, 0, 0), { itab: 4 })
      ],
      sheets: sheets("One", "Two", "Three")
    })
    expect(wb.definedNames).toEqual([
      { name: "Second", formula: "Two!$A$1", hidden: false, scope: 1 },
      { name: "First", formula: "One!$A$1", hidden: false, scope: 2 },
      { name: "Lost", formula: "Three!$A$1", hidden: false }
    ])
  })
})

/** The workbook stream with the globals BOF's version word replaced. */
const withVersion = (book: X.TestBook, version: number) => {
  const stream = X.workbookStream(book)
  stream.set(X.u16(version), 4)
  return stream
}

describe("xls globals: BOF version, sheets, SST", () => {
  it("BOF versions 0x0500-0x05FF are BIFF5, in a compound file or bare; others are UnsupportedFormat", async () => {
    const book: X.TestBook = { biff: 5, sheets: [{ name: "Five", records: [X.label5(0, 0, 0, [0x68, 0x69])] }] }
    for (const version of [0x0500, 0x0580, 0x05FF]) {
      const stream = withVersion(book, version)
      for (const bytes of [writeCfb([{ name: "Book", data: stream }]).bytes, stream]) {
        const exit = await Effect.runPromise(Effect.exit(Spreadsheet.read(bytes)))
        expect(Exit.isSuccess(exit) && exit.value.sheets.map((s) => s.name)).toEqual(["Five"])
        if (Exit.isSuccess(exit)) expect(await texts(exit.value)).toEqual(["hi"])
      }
    }
    for (const version of [0x0400, 0x04FF, 0x0601, 0x0700, 0x0000]) {
      const later = writeCfb([{ name: "Book", data: withVersion(book, version) }]).bytes
      expect(failureOf(await run(later))).toMatchObject({ _tag: "UnsupportedFormat", format: "xls", reason: `BIFF version 0x${version.toString(16)} in a BIFF5/8 BOF` })
    }
  })

  it("visibility comes from the low 2 bits of hsState only", async () => {
    const wb = await read({
      sheets: [0x00, 0x01, 0x02, 0x04, 0x05, 0x06, 0xFC, 0xFD, 0xFE].map((state, k) => ({ name: `S${k}`, records: [], state }))
    })
    expect(wb.sheets.map((s) => s.visibility)).toEqual(["visible", "hidden", "veryHidden", "visible", "hidden", "veryHidden", "visible", "hidden", "veryHidden"])
  })

  it("an SST in a BIFF5 workbook is skipped, whatever it holds", async () => {
    const bogus = X.record(0x00FC, [...X.u32(1e9), ...X.u32(1e9), 0xFF])
    const wb = await read({ biff: 5, globals: [bogus], sheets: [{ name: "x", records: [X.label5(0, 0, 0, [0x6F, 0x6B])] }] }, { maxSharedStrings: 10 })
    expect(await texts(wb)).toEqual(["ok"])
  })

  it("a short SST is MalformedFile; an SST with fewer strings than declared keeps those present", async () => {
    expect(await errorOf({ globals: [X.record(0x00FC, X.u32(1))], sheets: sheets("S") })).toMatchObject({ _tag: "MalformedFile", reason: "truncated SST" })
    const partial = X.record(0x00FC, [...X.u32(5), ...X.u32(5), ...X.xlString("one"), ...X.xlString("two"), 0x09])
    const wb = await read({ globals: [partial], sheets: [{ name: "S", records: [X.labelSst(0, 0, 0, 0), X.labelSst(0, 1, 0, 1)] }] })
    expect(await texts(wb)).toEqual(["one", "two"])
  })

  it("compressed (8-bit) BIFF8 characters 0x80-0x9F are Latin-1 code points, not windows-1252", async () => {
    const text = "A\u0080\u008A\u009F\u00E9" // all below 0x100: stored compressed, one byte each
    const wb = await read({ globals: [X.codepage(1252), X.sst([text])], sheets: [{ name: text, records: [X.labelSst(0, 0, 0, 0)] }] })
    expect(wb.sheets[0]!.name).toBe(text)
    expect(await texts(wb)).toEqual([text])
  })
})

describe("xls globals: text decoders and containers", () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it("an encoding the runtime lacks falls back to windows-1252", async () => {
    expect(textDecoder("x-no-such-encoding").encoding).toBe("windows-1252")
    expect(textDecoder("windows-1251").encoding).toBe("windows-1251")
    // a runtime without Shift-JIS (small ICU): CODEPAGE 932 still reads, as windows-1252
    const Real = globalThis.TextDecoder
    vi.stubGlobal("TextDecoder", class extends Real {
      constructor(label?: string, options?: TextDecoderOptions) {
        if (label === "shift_jis") throw new RangeError(`unknown encoding ${label}`)
        super(label, options)
      }
    })
    const wb = await read({ biff: 5, globals: [X.codepage(932)], sheets: [{ name: "x", name5: [0x82, 0xA0], records: [] }] })
    expect(wb.sheets[0]!.name).toBe("\u201A\u00A0")
  })

  it("CODEPAGE 28602 (no such ISO-8859 part) reads as windows-1252", async () => {
    const wb = await read({ biff: 5, globals: [X.codepage(28602)], sheets: [{ name: "x", name5: [0x80, 0x41], records: [X.label5(0, 0, 0, [0x80])] }] })
    expect(wb.sheets[0]!.name).toBe("\u20ACA")
    expect(await texts(wb)).toEqual(["\u20AC"])
  })

  it("a compound file whose workbook stream is BIFF2-4 is UnsupportedFormat naming the version", async () => {
    for (const v of [2, 3, 4]) {
      const biff = new Uint8Array(readFileSync(`test/corpus/generated/basic.biff${v}.xls`))
      for (const stream of ["Book", "Workbook"]) {
        const bytes = writeCfb([{ name: stream, data: biff }]).bytes
        for (const exit of [await run(bytes), await Effect.runPromise(Effect.exit(Spreadsheet.read(bytes)))]) {
          expect(failureOf(exit)).toMatchObject({ _tag: "UnsupportedFormat", format: `biff${v}` })
        }
      }
    }
  })
})

describe("xls globals: limits and string bounds", () => {
  it("Xls.read checks maxInputBytes before anything else", async () => {
    const bytes = X.xls({ sheets: sheets("S") })
    expect(failureOf(await run(bytes, { maxInputBytes: bytes.length - 1 }))).toMatchObject({ _tag: "LimitExceeded", limit: "maxInputBytes", max: bytes.length - 1, actual: bytes.length })
    expect(Exit.isSuccess(await run(bytes, { maxInputBytes: bytes.length }))).toBe(true)
  })

  it("an SST joined from CONTINUEs past maxDecompressedBytes is LimitExceeded", async () => {
    const strings = Array.from({ length: 60 }, (_, k) => `string ${k}`)
    const data = [...X.u32(60), ...X.u32(60), ...strings.flatMap(X.xlString)]
    const cut = 8 + strings.slice(0, 10).flatMap(X.xlString).length // at a string boundary
    const globals = [X.record(0x00FC, data.slice(0, cut)), X.record(CONTINUE, data.slice(cut))]
    const stream = X.workbookStream({ globals, sheets: [{ name: "S", records: [X.labelSst(0, 0, 0, 59)] }] })
    const total = data.length + 8
    expect(failureOf(await run(stream, { maxDecompressedBytes: total - 1 }))).toMatchObject({ _tag: "LimitExceeded", limit: "maxDecompressedBytes", max: total - 1 })
    const exit = await run(stream, { maxDecompressedBytes: total })
    expect(Exit.isSuccess(exit) && await texts(exit.value)).toEqual(["string 59"])
  })

  it("FORMAT codes and NAME names longer than maxStringLength are LimitExceeded", async () => {
    const limits = { maxStringLength: 20 }
    expect(await errorOf({ globals: [X.format(164, "0".repeat(21))], sheets: sheets("S") }, limits)).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
    expect(await errorOf({ biff: 5, globals: [X.format(164, "0".repeat(21), 5)], sheets: sheets("S") }, limits)).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
    expect(await errorOf({ globals: [X.name("N".repeat(21), ref3d(0x3A, 0, 0, 0))], sheets: sheets("S") }, limits)).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", max: 20, actual: 21 })
    const wb = await read({ globals: [X.format(164, "0".repeat(20)), X.xf(164), X.name("N".repeat(20), ref3d(0x3A, 0, 0, 0))], sheets: sheets("S") }, limits)
    expect(wb.definedNames.map((d) => d.name)).toEqual(["N".repeat(20)])
  })

  it("readString: strings end where their characters end, and running past the data is undefined", () => {
    const biff8: Text = { biff8: true, decoder: new TextDecoder("windows-1252"), maxLength: 100 }
    const biff5: Text = { ...biff8, biff8: false }
    const bytes = (...parts: ReadonlyArray<ReadonlyArray<number>>) => new Uint8Array(parts.flat())
    const pad = [9, 9, 9]
    // BIFF8 XLUnicodeString: u16 count, flags, characters
    expect(readString(bytes(pad, X.xlString("abc"), [7, 7]), 3, biff8, 2)).toEqual({ value: "abc", end: 9 })
    expect(readString(bytes(pad, X.xlString("añ\u0101")), 3, biff8, 2)).toEqual({ value: "añ\u0101", end: 12 })
    expect(readString(bytes(pad, X.xlString("abc").slice(0, 5)), 3, biff8, 2)).toBeUndefined()
    // BIFF8 ShortXLUnicodeString: u8 count, flags, characters
    expect(readString(bytes(pad, X.shortString("abc"), [7]), 3, biff8, 1)).toEqual({ value: "abc", end: 8 })
    expect(readString(bytes(pad, X.shortString("abc").slice(0, 4)), 3, biff8, 1)).toBeUndefined()
    expect(readString(bytes(pad, X.shortString("\u0101b").slice(0, 5)), 3, biff8, 1)).toBeUndefined()
    expect(readString(bytes(pad, [3]), 3, biff8, 1)).toBeUndefined()
    // BIFF5: u8 or u16 count, code-page bytes
    expect(readString(bytes(pad, X.string5([0x61, 0x80], 1), [7]), 3, biff5, 1)).toEqual({ value: "a\u20AC", end: 6 })
    expect(readString(bytes(pad, X.string5([0x61, 0x80], 2), [7]), 3, biff5, 2)).toEqual({ value: "a\u20AC", end: 7 })
    expect(readString(bytes(pad, X.string5([0x61, 0x80], 1).slice(0, 2)), 3, biff5, 1)).toBeUndefined()
    expect(readString(bytes(pad, X.string5([0x61, 0x80], 2).slice(0, 3)), 3, biff5, 2)).toBeUndefined()
    expect(readString(bytes(pad, [1]), 3, biff5, 2)).toBeUndefined()
    // past the data, and over the limit
    expect(readString(bytes(pad), 4, biff8, 2)).toBeUndefined()
    expect(readString(bytes(X.string5([0x61, 0x62], 1)), 0, { ...biff5, maxLength: 1 }, 1)).toBeNull()
  })
})

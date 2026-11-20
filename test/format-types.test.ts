/**
 * The format types: a workbook's format is one a reader reads, detection
 * reports any format, `UnsupportedFormat` names only what isn't read, and
 * `MalformedFile` and `Encrypted` name only the layers that can raise them.
 * The type assertions are checked by `tsc` (`pnpm check`); the runtime
 * tests pin the error schema's literal list to its type.
 */
import { assert, describe, expectTypeOf, it } from "@effect/vitest"
import { type Effect, Schema } from "effect"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as SpreadsheetError from "../src/SpreadsheetError.ts"
import * as Workbook from "../src/Workbook.ts"

type Readable = "csv" | "xlsx" | "xlsb" | "xls"
type Unreadable =
  | "ods"
  | "spreadsheetml"
  | "sylk"
  | "cfb"
  | "zip"
  | "html"
  | "xml"
  | "pdf"
  | "image"
  | "rtf"
  | "unknown"
type UnsupportedName = Unreadable | "biff2" | "biff3" | "biff4" | "xls"

/** Every name an `UnsupportedFormat` can carry, once: a missing or extra key doesn't compile. */
const UNSUPPORTED_NAMES: Record<UnsupportedName, true> = {
  ods: true,
  spreadsheetml: true,
  sylk: true,
  cfb: true,
  zip: true,
  html: true,
  xml: true,
  pdf: true,
  image: true,
  rtf: true,
  unknown: true,
  biff2: true,
  biff3: true,
  biff4: true,
  xls: true
}

type MalformedName = "csv" | "xlsx" | "xlsb" | "xls" | "zip" | "xml" | "cfb"
type EncryptedName = "xls" | "xlsx" | "zip"

/** Every name a `MalformedFile` can carry, once. */
const MALFORMED_NAMES: Record<MalformedName, true> = {
  csv: true,
  xlsx: true,
  xlsb: true,
  xls: true,
  zip: true,
  xml: true,
  cfb: true
}

/** Every name an `Encrypted` can carry, once. */
const ENCRYPTED_NAMES: Record<EncryptedName, true> = { xls: true, xlsx: true, zip: true }

/** A switch over a workbook's format that names only the readable formats. */
const readerOf = (format: Workbook.Workbook["format"]): string => {
  switch (format) {
    case "csv":
    case "xlsx":
    case "xlsb":
    case "xls":
      return format
    default:
      expectTypeOf(format).toBeNever()
      return format
  }
}

describe("format types", () => {
  it("a workbook's format is a readable format", () => {
    expectTypeOf<Workbook.ReadableFormat>().toEqualTypeOf<Readable>()
    expectTypeOf<Workbook.Workbook["format"]>().toEqualTypeOf<Readable>()
    expectTypeOf<Parameters<typeof Workbook.make>[0]["format"]>().toEqualTypeOf<Readable>()
    // @ts-expect-error a workbook is never read as a format without a reader
    const pdf: Workbook.Workbook["format"] = "pdf"
    // @ts-expect-error nor as a compound file without a workbook stream
    const cfb: Workbook.Workbook["format"] = "cfb"
    assert.deepStrictEqual([pdf, cfb, readerOf("xls")], ["pdf", "cfb", "xls"])
  })

  it("detection reports a readable or an unreadable format", () => {
    expectTypeOf<Workbook.UnreadableFormat>().toEqualTypeOf<Unreadable>()
    expectTypeOf<Workbook.Format>().toEqualTypeOf<Readable | Unreadable>()
    expectTypeOf<Spreadsheet.Format>().toEqualTypeOf<Workbook.Format>()
    expectTypeOf<Spreadsheet.ReadableFormat>().toEqualTypeOf<Workbook.ReadableFormat>()
    expectTypeOf<Spreadsheet.UnreadableFormat>().toEqualTypeOf<Workbook.UnreadableFormat>()
    expectTypeOf<Effect.Success<ReturnType<typeof Spreadsheet.detectFormat>>>().toEqualTypeOf<Readable | Unreadable>()
  })

  it("UnsupportedFormat names an unreadable format or an XLS version that isn't read", () => {
    expectTypeOf<SpreadsheetError.UnsupportedFormat["format"]>().toEqualTypeOf<UnsupportedName>()
    // @ts-expect-error CSV always has a reader
    const csv: SpreadsheetError.UnsupportedFormat["format"] = "csv"
    // @ts-expect-error so does an OOXML package
    const xlsx: SpreadsheetError.UnsupportedFormat["format"] = "xlsx"
    assert.deepStrictEqual([csv, xlsx], ["csv", "xlsx"])
  })

  it("UnsupportedFormat's schema accepts exactly the names of its type", () => {
    const literals: ReadonlyArray<string> = SpreadsheetError.UnsupportedFormat.fields.format.literals
    assert.deepStrictEqual([...literals].sort(), Object.keys(UNSUPPORTED_NAMES).sort())
    for (const format of literals) {
      const error = new SpreadsheetError.UnsupportedFormat({ format: format as UnsupportedName, reason: "r" })
      assert.isTrue(Schema.is(SpreadsheetError.SpreadsheetError)(error))
    }
    for (const format of ["csv", "xlsx", "xlsb", "biff5/8", ""]) {
      assert.throws(() => new SpreadsheetError.UnsupportedFormat({ format: format as UnsupportedName, reason: "r" }))
    }
  })

  it("MalformedFile names a reader's format or the container or XML layer under it", () => {
    expectTypeOf<SpreadsheetError.MalformedFile["format"]>().toEqualTypeOf<MalformedName>()
    // @ts-expect-error a format without a reader is never parsed far enough to be malformed
    const pdf: SpreadsheetError.MalformedFile["format"] = "pdf"
    // @ts-expect-error nor is an old BIFF version, which fails with UnsupportedFormat
    const biff4: SpreadsheetError.MalformedFile["format"] = "biff4"
    assert.deepStrictEqual([pdf, biff4], ["pdf", "biff4"])
  })

  it("Encrypted names an XLS workbook, an encrypted OOXML package or a zip entry", () => {
    expectTypeOf<SpreadsheetError.Encrypted["format"]>().toEqualTypeOf<EncryptedName>()
    // @ts-expect-error CSV has no encryption
    const csv: SpreadsheetError.Encrypted["format"] = "csv"
    // @ts-expect-error an encrypted XLSB package can't be told from an XLSX one, so it is named xlsx
    const xlsb: SpreadsheetError.Encrypted["format"] = "xlsb"
    assert.deepStrictEqual([csv, xlsb], ["csv", "xlsb"])
  })

  it("MalformedFile's and Encrypted's schemas accept exactly the names of their types", () => {
    const malformed: ReadonlyArray<string> = SpreadsheetError.MalformedFile.fields.format.literals
    assert.deepStrictEqual([...malformed].sort(), Object.keys(MALFORMED_NAMES).sort())
    for (const format of malformed) {
      const error = new SpreadsheetError.MalformedFile({ format: format as MalformedName, reason: "r" })
      assert.isTrue(Schema.is(SpreadsheetError.SpreadsheetError)(error))
    }
    for (const format of ["ods", "pdf", "unknown", "biff4", ""]) {
      assert.throws(() => new SpreadsheetError.MalformedFile({ format: format as MalformedName, reason: "r" }))
    }
    const encrypted: ReadonlyArray<string> = SpreadsheetError.Encrypted.fields.format.literals
    assert.deepStrictEqual([...encrypted].sort(), Object.keys(ENCRYPTED_NAMES).sort())
    for (const format of encrypted) {
      const error = new SpreadsheetError.Encrypted({ format: format as EncryptedName })
      assert.isTrue(Schema.is(SpreadsheetError.SpreadsheetError)(error))
    }
    for (const format of ["csv", "xlsb", "cfb", ""]) {
      assert.throws(() => new SpreadsheetError.Encrypted({ format: format as EncryptedName }))
    }
  })
})

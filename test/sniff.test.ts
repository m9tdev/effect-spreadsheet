import { assert, describe, it } from "@effect/vitest"
import { Effect } from "effect"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import { enc } from "./helpers/run.ts"
import { makeZip } from "./helpers/zip.ts"

const detect = (bytes: Uint8Array) => Effect.runPromise(Effect.exit(Spreadsheet.detectFormat(bytes)))
const formatOf = async (bytes: Uint8Array) => {
  const exit = await detect(bytes)
  return exit._tag === "Success" ? exit.value : `error:${exit.cause}`
}

const CFB = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, ...new Uint8Array(504)])
const XLSX = makeZip([{ name: "[Content_Types].xml" }, { name: "xl/workbook.xml" }, { name: "xl/worksheets/sheet1.xml" }])
const XLSB = makeZip([{ name: "[Content_Types].xml" }, { name: "xl/workbook.bin" }])

describe("format detection", () => {
  it("zip containers by their entries", async () => {
    assert.strictEqual(await formatOf(XLSX), "xlsx")
    assert.strictEqual(await formatOf(XLSB), "xlsb")
    assert.strictEqual(
      await formatOf(makeZip([{ name: "mimetype", data: enc("application/vnd.oasis.opendocument.spreadsheet") }, { name: "content.xml" }])),
      "ods"
    )
    assert.strictEqual(
      await formatOf(makeZip([{ name: "mimetype", data: enc("application/vnd.oasis.opendocument.text") }, { name: "content.xml" }])),
      "zip"
    )
    assert.strictEqual(await formatOf(makeZip([{ name: "hello.txt", data: enc("hi") }])), "zip")
  })

  it("a zip with an archive comment still has its directory found", async () => {
    const zip = makeZip([{ name: "xl/workbook.xml" }])
    const withComment = new Uint8Array(zip.length + 1000)
    withComment.set(zip)
    new DataView(withComment.buffer).setUint16(zip.length - 2, 1000, true)
    assert.strictEqual(await formatOf(withComment), "xlsx")
  })

  it("truncated or corrupt zip is MalformedFile, too many entries is LimitExceeded", async () => {
    const exit = await detect(XLSX.subarray(0, XLSX.length - 10))
    assert.strictEqual(exit._tag === "Failure" && String(exit.cause).includes("MalformedFile"), true)
    const many = makeZip(Array.from({ length: 20 }, (_, i) => ({ name: `f${i}` })))
    const limited = await Effect.runPromise(Effect.flip(Spreadsheet.detectFormat(many, { limits: { maxZipEntries: 10 } })))
    assert.strictEqual(limited._tag === "LimitExceeded" && limited.limit, "maxZipEntries")
    // central directory offset pointing past the end
    const bad = XLSX.slice()
    new DataView(bad.buffer).setUint32(bad.length - 6, 0xFFFFFF00, true)
    const badExit = await Effect.runPromise(Effect.flip(Spreadsheet.detectFormat(bad)))
    assert.strictEqual(badExit._tag, "MalformedFile")
  })

  it("CFB, raw BIFF, SpreadsheetML, SYLK, HTML, XML", async () => {
    assert.strictEqual(await formatOf(CFB), "cfb")
    assert.strictEqual(await formatOf(new Uint8Array([0x09, 0x00, 0x04, 0x00, 0x02, 0x00, 0x10, 0x00])), "xls")
    assert.strictEqual(await formatOf(new Uint8Array([0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x10, 0x00])), "xls")
    const xml = `<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"></Workbook>`
    assert.strictEqual(await formatOf(enc(xml)), "spreadsheetml")
    assert.strictEqual(await formatOf(new Uint8Array([0xEF, 0xBB, 0xBF, ...enc("  \n" + xml)])), "spreadsheetml")
    assert.strictEqual(await formatOf(enc("ID;PWXL;N;E\r\nC;Y1;X1;K1\r\nE\r\n")), "sylk")
    assert.strictEqual(await formatOf(enc("<html><body><table></table></body></html>")), "html")
    assert.strictEqual(await formatOf(enc("<?xml version=\"1.0\"?><foo/>")), "xml")
  })

  it("raw BIFF needs a real BOF: UTF-16LE text starting with tabs is CSV", async () => {
    const utf16le = (text: string) => new Uint8Array([...text].flatMap((c) => [c.charCodeAt(0), 0]))
    assert.strictEqual(await formatOf(utf16le("\t\tname\tx\n1\t2\t3\n")), "csv")
    assert.strictEqual(await formatOf(utf16le("\t\na\tb\n")), "csv")
    // BIFF5/8 BOF: version 0x0600, workbook globals
    assert.strictEqual(await formatOf(new Uint8Array([0x09, 0x08, 0x10, 0x00, 0x00, 0x06, 0x05, 0x00, ...new Uint8Array(12)])), "xls")
    // right type, wrong length or unknown data type
    assert.notStrictEqual(await formatOf(new Uint8Array([0x09, 0x00, 0x06, 0x00, 0x02, 0x00, 0x10, 0x00])), "xls")
    assert.notStrictEqual(await formatOf(new Uint8Array([0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x33, 0x00])), "xls")
    assert.notStrictEqual(await formatOf(new Uint8Array([0x09, 0x08, 0x10, 0x00, 0x00, 0x07, 0x05, 0x00, ...new Uint8Array(12)])), "xls")
  })

  it("a CSV whose first field starts with < is CSV; markup is html/xml", async () => {
    assert.strictEqual(await formatOf(enc("<b>,x\n1,2\n")), "csv")
    assert.strictEqual(await formatOf(enc("<b>name</b>;<i>x</i>\n")), "csv")
    assert.strictEqual(await formatOf(enc("<5,x\n")), "csv")
    assert.strictEqual(await formatOf(enc("<a\n")), "csv")
    assert.strictEqual(await formatOf(enc("<table border=1>\n<tr><td>a, b</td></tr></table>")), "html")
    assert.strictEqual(await formatOf(enc("<foo/>")), "xml")
    assert.strictEqual(await formatOf(enc("<!-- c -->,x")), "xml")
    const wb = await Effect.runPromise(Spreadsheet.read(enc("<b>,x\n")))
    assert.strictEqual(wb.format, "csv")
  })

  it("a semicolon CSV starting with ID;P is CSV, not SYLK", async () => {
    assert.strictEqual(await formatOf(enc("ID;Price;Name\n1;2;x\n")), "csv")
  })

  it("PDF, images, unknown binary", async () => {
    assert.strictEqual(await formatOf(enc("%PDF-1.7\n")), "pdf")
    assert.strictEqual(await formatOf(new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A])), "image")
    assert.strictEqual(await formatOf(new Uint8Array([0x01, 0x02, 0x00, 0x00, 0x05, 0x00, 0x07])), "unknown")
  })

  it("text and empty input are CSV", async () => {
    assert.strictEqual(await formatOf(new Uint8Array(0)), "csv")
    assert.strictEqual(await formatOf(enc("a,b\n")), "csv")
    assert.strictEqual(await formatOf(new Uint8Array([0xFF, 0xFE, 0x61, 0x00, 0x2C, 0x00])), "csv")
  })
})

describe("Spreadsheet.read dispatch", () => {
  it.effect("unimplemented formats fail with UnsupportedFormat naming the format", () =>
    Effect.gen(function*() {
      for (const [bytes, format] of [[XLSX, "xlsx"], [XLSB, "xlsb"], [CFB, "cfb"], [enc("%PDF-1.4"), "pdf"]] as const) {
        const error = yield* Effect.flip(Spreadsheet.read(bytes))
        assert.strictEqual(error._tag, "UnsupportedFormat")
        if (error._tag === "UnsupportedFormat") assert.strictEqual(error.format, format)
      }
    }))

  it.effect("detection uses content only: there is no file-name input, CFB/ZIP bytes are never read as CSV", () =>
    Effect.gen(function*() {
      // A caller holding "report.csv" can only pass the bytes (plus options
      // that carry no name); the bytes decide.
      const options = { csv: { sheetName: "report.csv" } } as const
      for (const bytes of [CFB, XLSX, XLSB]) {
        const error = yield* Effect.flip(Spreadsheet.read(bytes, options))
        assert.strictEqual(error._tag, "UnsupportedFormat")
        if (error._tag === "UnsupportedFormat") assert.notStrictEqual(error.format, "csv")
      }
      // and text named like a workbook is still CSV
      const wb = yield* Spreadsheet.read(enc("a,b\n"), { csv: { sheetName: "book.xlsx" } })
      assert.strictEqual(wb.format, "csv")
    }))

  it.effect("maxInputBytes is checked before detection", () =>
    Effect.gen(function*() {
      const error = yield* Effect.flip(Spreadsheet.read(XLSX, { limits: { maxInputBytes: 10 } }))
      assert.strictEqual(error._tag === "LimitExceeded" && error.limit, "maxInputBytes")
    }))
})

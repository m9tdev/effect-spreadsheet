import { assert, describe, expect, it } from "@effect/vitest"
import { Effect, Stream } from "effect"
import * as Cell from "../src/Cell.ts"
import * as Csv from "../src/csv/index.ts"
import * as Spreadsheet from "../src/Spreadsheet.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Workbook from "../src/Workbook.ts"
import { csvAoa, csvExit, csvRows, enc, failureOf, toAoa } from "./helpers/run.ts"

describe("csv: structure", () => {
  it("empty input gives no rows", async () => {
    assert.deepStrictEqual(await csvAoa(""), [])
  })

  it("only a BOM gives no rows (UTF-8, UTF-16LE, UTF-16BE)", async () => {
    assert.deepStrictEqual(await csvAoa(new Uint8Array([0xEF, 0xBB, 0xBF])), [])
    assert.deepStrictEqual(await csvAoa(new Uint8Array([0xFF, 0xFE])), [])
    assert.deepStrictEqual(await csvAoa(new Uint8Array([0xFE, 0xFF])), [])
  })

  it("basic rows, trailing line break does not add a row", async () => {
    assert.deepStrictEqual(await csvAoa("a,b,c\n1,2,3\n"), [["a", "b", "c"], ["1", "2", "3"]])
    assert.deepStrictEqual(await csvAoa("a,b,c\n1,2,3"), [["a", "b", "c"], ["1", "2", "3"]])
  })

  it("LF, CRLF and CR line breaks", async () => {
    const expected = [["a", "b"], ["c", "d"], ["e", "f"]]
    assert.deepStrictEqual(await csvAoa("a,b\nc,d\ne,f"), expected)
    assert.deepStrictEqual(await csvAoa("a,b\r\nc,d\r\ne,f\r\n"), expected)
    assert.deepStrictEqual(await csvAoa("a,b\rc,d\re,f\r"), expected)
    assert.deepStrictEqual(await csvAoa("a,b\r\nc,d\re,f\n"), expected)
  })

  it("blank lines are rows without cells; indexes count every record", async () => {
    const rows = await Effect.runPromise(csvRows("a\n\n\nb\n"))
    assert.deepStrictEqual(rows.map((r) => [r.index, r.cells.length]), [[0, 1], [1, 0], [2, 0], [3, 1]])
    assert.isTrue(rows.every((r) => r.hidden === false))
  })

  it("ragged rows and empty fields (sparse cells keep their column)", async () => {
    const rows = await Effect.runPromise(csvRows("a,,c\n,x\n1,2,3,4,5\n,,\n"))
    assert.deepStrictEqual(rows.map((r) => r.cells.map((c) => [c.column, Cell.value(c)])), [
      [[0, "a"], [2, "c"]],
      [[1, "x"]],
      [[0, "1"], [1, "2"], [2, "3"], [3, "4"], [4, "5"]],
      []
    ])
  })

  it("RFC 4180 quoting: escaped quotes, delimiters and line breaks inside quotes", async () => {
    assert.deepStrictEqual(await csvAoa('"a,b","c""d","e\nf","g\r\nh",""\n"x"'), [
      ["a,b", 'c"d', "e\nf", "g\r\nh"],
      ["x"]
    ])
    assert.deepStrictEqual(await csvAoa('""""'), [['"']])
  })

  it("quote inside an unquoted field is literal", async () => {
    assert.deepStrictEqual(await csvAoa('ab"c,d"\n'), [['ab"c', 'd"']])
  })

  it("text after a closing quote is appended (lenient) or rejected (strict)", async () => {
    assert.deepStrictEqual(await csvAoa('"ab"c,d'), [["abc", "d"]])
    const exit = await csvExit('"ab"c,d', { strict: true })
    const error = failureOf(exit)
    assert.strictEqual(error?._tag, "MalformedFile")
    if (error?._tag === "MalformedFile") assert.strictEqual(error.offset, 4)
  })

  it("unterminated quote at EOF keeps the rest (lenient) or fails (strict)", async () => {
    assert.deepStrictEqual(await csvAoa('a,"b\nc,d'), [["a", "b\nc,d"]])
    const error = failureOf(await csvExit('a,"b\nc,d', { strict: true }))
    assert.strictEqual(error?._tag, "MalformedFile")
    if (error?._tag === "MalformedFile") assert.match(error.reason, /unterminated/)
    // a properly closed quote at EOF is fine in strict mode
    assert.deepStrictEqual(await csvAoa('a,"b"', { strict: true }), [["a", "b"]])
  })

  it("whitespace is kept", async () => {
    assert.deepStrictEqual(await csvAoa(" a , b \n"), [[" a ", " b "]])
  })
})

describe("csv: delimiters", () => {
  it("explicit delimiter", async () => {
    assert.deepStrictEqual(await csvAoa("a;b,c\n", { delimiter: ";" }), [["a", "b,c"]])
  })

  it("sniffs , ; tab |", async () => {
    assert.deepStrictEqual(await csvAoa("a;b;c\n1;2;3\n"), [["a", "b", "c"], ["1", "2", "3"]])
    assert.deepStrictEqual(await csvAoa("a\tb\n1\t2\n"), [["a", "b"], ["1", "2"]])
    assert.deepStrictEqual(await csvAoa("a|b\n1|2\n"), [["a", "b"], ["1", "2"]])
    assert.deepStrictEqual(await csvAoa("a,b\n1,2\n"), [["a", "b"], ["1", "2"]])
  })

  it("sniffing prefers a delimiter with a consistent count per line", async () => {
    // commas inside values (European decimals) but 2 semicolons on every line
    assert.deepStrictEqual(await csvAoa("x;y;z\n1,5;2,25,1;3\n4;5;6\n"), [["x", "y", "z"], ["1,5", "2,25,1", "3"], ["4", "5", "6"]])
  })

  it("sniffing ignores delimiters inside quotes", async () => {
    assert.deepStrictEqual(await csvAoa('"a;b;c";"d"\n"e;f";"g"\n'), [["a;b;c", "d"], ["e;f", "g"]])
  })

  it("sniffing skips quoted delimiters (only `;` is outside quotes here)", async () => {
    assert.deepStrictEqual(await csvAoa('"a,b,c";d\n"e,f,g";h\n'), [["a,b,c", "d"], ["e,f,g", "h"]])
  })

  it("sniffing ignores a line cut off at the end of the 16 Ki sample", async () => {
    // 297-char lines with 1 `,` and 2 `;`: the sample ends 49 chars into a
    // line holding only the `,`, which would make `;` look inconsistent.
    const line = "a," + "x".repeat(290) + ";c;d\n"
    expect(line.length).toBe(297)
    expect(16 * 1024 % 297).toBe(49)
    const rows = await csvAoa(line.repeat(60))
    assert.deepStrictEqual(rows[0], ["a," + "x".repeat(290), "c", "d"])
  })

  it("no candidate: comma", async () => {
    assert.deepStrictEqual(await csvAoa("abc\ndef\n"), [["abc"], ["def"]])
  })

  it("Excel sep= line sets the delimiter and is not a row", async () => {
    assert.deepStrictEqual(await csvAoa("sep=;\r\na;b,c\n"), [["a", "b,c"]])
    assert.deepStrictEqual(await csvAoa("sep=|\na|b\n"), [["a", "b"]])
    // an explicit delimiter wins and the line stays data
    assert.deepStrictEqual(await csvAoa("sep=;\na,b\n", { delimiter: "," }), [["sep=;"], ["a", "b"]])
  })

  it("invalid delimiter option is an InvalidOption error", async () => {
    for (const delimiter of ['"', ";;", "", "\n", "\r"]) {
      const error = failureOf(await csvExit("a", { delimiter }))
      assert.deepStrictEqual(error?._tag === "InvalidOption" && error.option, "delimiter")
    }
    const atRead = await Effect.runPromise(Effect.flip(Csv.read(enc("a"), { delimiter: ";;" })))
    assert.deepStrictEqual(atRead._tag === "InvalidOption" && atRead.option, "delimiter")
  })
})

describe("csv: encodings", () => {
  const text = "név,ár\nÉva,€5\n"
  const expected = [["név", "ár"], ["Éva", "€5"]]
  const utf16 = (s: string, be: boolean, bom: boolean) => {
    const out = new Uint8Array((s.length + (bom ? 1 : 0)) * 2)
    const all = (bom ? "\uFEFF" : "") + s
    for (let i = 0; i < all.length; i++) {
      const c = all.charCodeAt(i)
      out[2 * i + (be ? 1 : 0)] = c & 0xFF
      out[2 * i + (be ? 0 : 1)] = c >> 8
    }
    return out
  }

  it("UTF-8 with and without BOM", async () => {
    assert.deepStrictEqual(await csvAoa(text), expected)
    assert.deepStrictEqual(await csvAoa(new Uint8Array([0xEF, 0xBB, 0xBF, ...enc(text)])), expected)
  })

  it("UTF-16LE/BE with BOM, and without BOM (NUL pattern)", async () => {
    assert.deepStrictEqual(await csvAoa(utf16(text, false, true)), expected)
    assert.deepStrictEqual(await csvAoa(utf16(text, true, true)), expected)
    assert.deepStrictEqual(await csvAoa(utf16(text, false, false)), expected)
    assert.deepStrictEqual(await csvAoa(utf16(text, true, false)), expected)
  })

  it("invalid UTF-8 falls back to windows-1252", async () => {
    // "név,€" in windows-1252: é = 0xE9, € = 0x80
    const bytes = new Uint8Array([0x6E, 0xE9, 0x76, 0x2C, 0x80, 0x0A])
    assert.deepStrictEqual(await csvAoa(bytes), [["név", "€"]])
  })

  it("explicit encoding", async () => {
    const bytes = new Uint8Array([0x6E, 0xE9, 0x76])
    assert.deepStrictEqual(await csvAoa(bytes, { encoding: "windows-1252" }), [["név"]])
    assert.deepStrictEqual(await csvAoa(new Uint8Array([0xEF, 0xBB, 0xBF, 0x61]), { encoding: "utf-8" }), [["a"]])
    assert.deepStrictEqual(await csvAoa(utf16("a,b", false, true), { encoding: "utf-16le" }), [["a", "b"]])
  })

  it("unknown encoding label is an InvalidOption error (when streaming and at read)", async () => {
    const error = failureOf(await csvExit("a", { encoding: "klingon" }))
    assert.deepStrictEqual(error?._tag === "InvalidOption" && error.option, "encoding")
    const atRead = await Effect.runPromise(Effect.flip(Csv.read(enc("a"), { encoding: "klingon" })))
    assert.deepStrictEqual(atRead._tag === "InvalidOption" && atRead.option, "encoding")
  })

  it("encoding is decided on the first 64 KiB only: an invalid byte near 70 000 is U+FFFD, however chunked", async () => {
    const head = enc("a,é\n".repeat(14_000)) // 70 000 bytes of valid UTF-8
    const bytes = new Uint8Array([...head, 0xFF, ...enc(",x\n")])
    expect(bytes.length).toBeGreaterThan(70_000)
    const expected = [...Array.from({ length: 14_000 }, () => ["a", "é"]), ["\uFFFD", "x"]]
    expect(await csvAoa(bytes)).toEqual(expected)
    const twoChunks = Stream.make(bytes.subarray(0, 40 * 1024), bytes.subarray(40 * 1024))
    expect(await csvAoa(twoChunks)).toEqual(expected)
  })

  it("truncated multibyte character at EOF: U+FFFD in UTF-8, windows-1252 if it makes the whole input invalid", async () => {
    const truncated = new Uint8Array([0x61, 0x2C, 0xE2, 0x82])
    assert.deepStrictEqual(await csvAoa(truncated, { encoding: "utf-8" }), [["a", "\uFFFD"]])
    assert.deepStrictEqual(await csvAoa(truncated), [["a", "\u00E2\u201A"]])
    // past the 64 KiB detection window the decision is UTF-8 already
    const long = new Uint8Array([...enc("é".repeat(40_000) + ","), 0xE2, 0x82])
    const rows = await csvAoa(long)
    assert.strictEqual(rows[0]![1], "\uFFFD")
  })
})

describe("csv: type inference", () => {
  it("off by default: all text", async () => {
    const rows = await Effect.runPromise(csvRows("1,TRUE,#N/A,x\n"))
    assert.deepStrictEqual(rows[0]!.cells.map((c) => c._tag), ["Text", "Text", "Text", "Text"])
  })

  it("numbers, booleans, errors with inferTypes", async () => {
    const rows = await Effect.runPromise(
      csvRows("1,-2.5,.5,1e3,+7,0,0.25,TRUE,false,True,#N/A,#DIV/0!,x,\"3\"\n", { inferTypes: true })
    )
    assert.deepStrictEqual(rows[0]!.cells.map((c) => [c._tag, Cell.value(c)]), [
      ["Number", 1],
      ["Number", -2.5],
      ["Number", 0.5],
      ["Number", 1000],
      ["Number", 7],
      ["Number", 0],
      ["Number", 0.25],
      ["Boolean", true],
      ["Boolean", false],
      ["Boolean", true],
      ["Error", "#N/A"],
      ["Error", "#DIV/0!"],
      ["Text", "x"],
      ["Number", 3]
    ])
  })

  it("negative zero is inferred as 0", async () => {
    const rows = await Effect.runPromise(csvRows("-0,-0.0,-0e5,+0\n", { inferTypes: true }))
    for (const cell of rows[0]!.cells) assert.isTrue(Object.is(Cell.value(cell), 0), String(Cell.value(cell)))
  })

  it("keeps as text: leading zeros, long digit strings, separators, dates, formulas, padding", async () => {
    const values = ["007", "1234567890123456", "1,000", "$5", "50%", "2024-01-02", "=SUM(A1)", " 1", "1 ", "1e999", "-", "."]
    const rows = await Effect.runPromise(csvRows(values.map((v) => `"${v}"`).join(",") + "\n", { inferTypes: true }))
    assert.deepStrictEqual(rows[0]!.cells.map((c) => [c._tag, Cell.value(c)]), values.map((v) => ["Text", v]))
  })
})

describe("csv: limits", () => {
  it("maxColumns: a huge single row fails with LimitExceeded", async () => {
    const row = Array.from({ length: 20_000 }, (_, i) => String(i)).join(",")
    const error = failureOf(await csvExit(row))
    assert.strictEqual(error?._tag, "LimitExceeded")
    if (error?._tag === "LimitExceeded") {
      assert.strictEqual(error.limit, "maxColumns")
      assert.strictEqual(error.max, 16_384)
      assert.strictEqual(error.actual, 16_385)
    }
    // empty fields count as columns too
    const empties = failureOf(await csvExit(",".repeat(100), { limits: { maxColumns: 50 } }))
    assert.strictEqual(empties?._tag === "LimitExceeded" && empties.limit, "maxColumns")
  })

  it("maxStringLength: a 1M-char field fails by default and passes when raised", async () => {
    const big = "x".repeat(1_000_000)
    const error = failureOf(await csvExit(`a,${big},b\n`))
    assert.strictEqual(error?._tag === "LimitExceeded" && error.limit, "maxStringLength")
    const quoted = failureOf(await csvExit(`"${big}"`))
    assert.strictEqual(quoted?._tag === "LimitExceeded" && quoted.limit, "maxStringLength")
    const rows = await Effect.runPromise(csvRows(`a,${big},b\n`, { limits: { maxStringLength: 2_000_000 } }))
    assert.strictEqual((Cell.value(rows[0]!.cells[1]!) as string).length, 1_000_000)
  })

  it("maxRows and maxCellsPerSheet", async () => {
    const rowsError = failureOf(await csvExit("a\nb\nc\n", { limits: { maxRows: 2 } }))
    assert.deepStrictEqual(
      rowsError?._tag === "LimitExceeded" ? [rowsError.limit, rowsError.max, rowsError.actual] : undefined,
      ["maxRows", 2, 3]
    )
    assert.strictEqual((await csvAoa("a\nb\n", { limits: { maxRows: 2 } })).length, 2)
    const cellsError = failureOf(await csvExit("a,b\nc,d\n", { limits: { maxCellsPerSheet: 3 } }))
    assert.strictEqual(cellsError?._tag === "LimitExceeded" && cellsError.limit, "maxCellsPerSheet")
    // empty fields don't count as cells
    assert.strictEqual((await csvAoa("a,,,\n,,b\n", { limits: { maxCellsPerSheet: 2 } })).length, 2)
  })

  it("maxInputBytes for arrays (at read) and streams (while reading)", async () => {
    const wbError = await Effect.runPromise(Effect.flip(Csv.read(enc("abcdef"), { limits: { maxInputBytes: 3 } })))
    assert.strictEqual(wbError._tag === "LimitExceeded" && wbError.limit, "maxInputBytes")
    const stream = Stream.fromIterable([enc("ab\n"), enc("cd\n"), enc("ef\n")])
    const error = failureOf(await csvExit(stream, { limits: { maxInputBytes: 5 } }))
    assert.strictEqual(error?._tag === "LimitExceeded" && error.limit, "maxInputBytes")
  })

  it.effect("limits come from the layer, per-call overrides win", () =>
    Effect.gen(function*() {
      const e = yield* Effect.flip(csvRows("a,b,c"))
      assert.strictEqual(e._tag === "LimitExceeded" && e.max, 2)
      const ok = yield* csvRows("a,b,c", { limits: { maxColumns: 3 } })
      assert.strictEqual(ok.length, 1)
    }).pipe(Effect.provide(SpreadsheetLimits.layer({ maxColumns: 2 }))))
})

describe("csv: streaming", () => {
  it.effect("byte stream input with chunks splitting a multibyte char, a quote pair and CRLF", () =>
    Effect.gen(function*() {
      const bytes = enc('é,"a""b"\r\n€,x\r\n')
      // é = C3 A9; cut inside it, inside "", and between CR and LF
      const cuts = [1, 6, 7, 11, 12]
      const chunks: Array<Uint8Array> = []
      let last = 0
      for (const c of cuts) {
        chunks.push(bytes.subarray(last, c))
        last = c
      }
      chunks.push(bytes.subarray(last))
      const rows = yield* csvRows(Stream.fromIterable(chunks))
      assert.deepStrictEqual(toAoa(rows), [["é", 'a"b'], ["€", "x"]])
    }))

  it.effect("rows come out in several chunks for large input", () =>
    Effect.gen(function*() {
      const line = "abcdefgh,ijklmnop,qrstuvwx\n"
      const bytes = enc(line.repeat(20_000)) // ~540 KB
      const chunks = yield* Stream.runCollect(Csv.rows(bytes).pipe(Stream.chunks))
      const sizes = Array.from(chunks).map((c) => c.length)
      assert.isAbove(sizes.length, 5)
      assert.strictEqual(sizes.reduce((a, b) => a + b, 0), 20_000)
    }))

  it("the event loop runs while a large input is parsed", async () => {
    const bytes = enc("abcdefgh,ijklmnop,qrstuvwx\n".repeat(100_000)) // ~2.7 MB
    let ticks = 0
    let running = true
    const tick = () => {
      if (!running) return
      ticks++
      setImmediate(tick)
    }
    setImmediate(tick)
    const count = await Effect.runPromise(Stream.runCount(Csv.rows(bytes)))
    running = false
    assert.strictEqual(count, 100_000)
    assert.isAbove(ticks, 10)
  })

  it.effect("the error of the input stream passes through", () =>
    Effect.gen(function*() {
      const input = Stream.concat(Stream.make(enc("a\n")), Stream.fail("boom" as const))
      const error = yield* Effect.flip(Stream.runCollect(Csv.rows(input)))
      assert.strictEqual(error, "boom")
    }))

  it.effect("rows of a Uint8Array can be streamed twice", () =>
    Effect.gen(function*() {
      const s = Csv.rows(enc("a,b\n"))
      assert.deepStrictEqual(toAoa(Array.from(yield* Stream.runCollect(s))), toAoa(Array.from(yield* Stream.runCollect(s))))
    }))
})

describe("csv: workbook", () => {
  it.effect("read gives one sheet, rows by name or index", () =>
    Effect.gen(function*() {
      const wb = yield* Csv.read(enc("a,b\n1,2\n"), { sheetName: "Data", inferTypes: true })
      assert.strictEqual(wb.format, "csv")
      assert.deepStrictEqual(wb.sheets, [{ name: "Data", index: 0, visibility: "visible" }])
      const rows = Array.from(yield* Stream.runCollect(Workbook.rows(wb, "Data")))
      assert.deepStrictEqual(toAoa(rows), [["a", "b"], [1, 2]])
      const again = Array.from(yield* Stream.runCollect(Workbook.rows(wb, 0, { limits: { maxRows: 5 } })))
      assert.deepStrictEqual(toAoa(again), toAoa(rows))
      const e = yield* Effect.flip(Stream.runCollect(Workbook.rows(wb, 0, { limits: { maxRows: 1 } })))
      assert.strictEqual(e._tag === "LimitExceeded" && e.limit, "maxRows")
    }))

  it.effect("limits from the layer at read time stay with the workbook (rows streamed outside that scope)", () =>
    Effect.gen(function*() {
      const bytes = enc("1\n2\n3\n4\n5\n6\n7\n8\n9\n10\n")
      for (const open of [Csv.read(bytes), Spreadsheet.read(bytes)]) {
        const wb = yield* open.pipe(Effect.provide(SpreadsheetLimits.layer({ maxRows: 3 })))
        assert.strictEqual(wb.limits.maxRows, 3)
        const error = yield* Effect.flip(Stream.runCollect(Workbook.rows(wb, 0)))
        assert.deepStrictEqual(error._tag === "LimitExceeded" && [error.limit, error.max], ["maxRows", 3])
        // a different layer at streaming time doesn't change them
        const still = yield* Effect.flip(Stream.runCollect(Workbook.rows(wb, 0))).pipe(
          Effect.provide(SpreadsheetLimits.layer({ maxRows: 100 }))
        )
        assert.strictEqual(still._tag === "LimitExceeded" && still.max, 3)
        // rows options apply on top
        const raised = yield* Stream.runCollect(Workbook.rows(wb, 0, { limits: { maxRows: 10 } }))
        assert.strictEqual(Array.from(raised).length, 10)
      }
      // read-call overrides are kept too
      const wb = yield* Csv.read(bytes, { limits: { maxColumns: 1, maxRows: 4 } })
      assert.deepStrictEqual([wb.limits.maxColumns, wb.limits.maxRows], [1, 4])
    }))

  it.effect("Spreadsheet.read dispatches text to CSV with csv options", () =>
    Effect.gen(function*() {
      const wb = yield* Spreadsheet.read(enc("x;1\n"), { csv: { inferTypes: true } })
      assert.strictEqual(wb.format, "csv")
      const rows = Array.from(yield* Stream.runCollect(Workbook.rows(wb, 0)))
      assert.deepStrictEqual(toAoa(rows), [["x", 1]])
      const viaReadCsv = yield* Spreadsheet.readCsv(enc("a"))
      assert.strictEqual(viaReadCsv.sheets[0]!.name, "Sheet1")
    }))
})

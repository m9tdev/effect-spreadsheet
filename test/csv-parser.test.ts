import { assert, describe, it } from "@effect/vitest"
import * as Cell from "../src/Cell.ts"
import { CsvParser } from "../src/internal/csv/parser.ts"
import { SNIFF_CHARS } from "../src/internal/csv/sniff.ts"
import type { SpreadsheetError } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import type { Row } from "../src/Workbook.ts"

const parser = (options: { delimiter?: string; strict?: boolean } = {}) =>
  new CsvParser({
    delimiter: options.delimiter,
    inferTypes: false,
    strict: options.strict ?? false,
    limits: SpreadsheetLimits.defaults
  })

const aoa = (rows: ReadonlyArray<Row>): Array<Array<unknown>> =>
  rows.map((row) => {
    const out: Array<unknown> = []
    for (const cell of row.cells) out[cell.column] = Cell.value(cell)
    return Array.from(out)
  })

const isError = (r: Array<Row> | SpreadsheetError): r is SpreadsheetError => !Array.isArray(r)

/** The rows of a feed that must not fail. */
const rowsOf = (r: Array<Row> | SpreadsheetError): Array<Row> => {
  if (isError(r)) return assert.fail(`unexpected ${r._tag}`)
  return r
}

/** The offset of a MalformedFile error, from a feed that must fail. */
const malformedOffset = (r: Array<Row> | SpreadsheetError): number | undefined => {
  if (!isError(r) || r._tag !== "MalformedFile") return assert.fail("expected a MalformedFile error")
  return r.offset
}

/**
 * Feeds the pieces in order (the last one as final) the way the reader does:
 * stop at the first error, after taking any rows that came before it.
 */
const feedAll = (pieces: ReadonlyArray<string>, options: { delimiter?: string; strict?: boolean } = {}) => {
  const p = parser(options)
  const rows: Array<Row> = []
  for (let k = 0; k < pieces.length; k++) {
    const r = p.feed(pieces[k]!, k === pieces.length - 1)
    if (isError(r)) return { rows: aoa(rows), error: r._tag }
    rows.push(...r)
    if (p.error !== undefined) return { rows: aoa(rows), error: p.error._tag }
  }
  return { rows: aoa(rows), error: undefined }
}

describe("csv parser: tokenizer states across pieces", () => {
  // Each prefix, fed as one piece, leaves the tokenizer in the named state.
  const prefixes: ReadonlyArray<[string, string]> = [
    ["start of input", ""],
    ["start of a record after LF", "a\n"],
    ["start of a record after CR", "a\r"],
    ["start of a field after a delimiter", "a,"],
    ["inside an unquoted field", "a,b"],
    ["inside a quoted field", 'a,"b'],
    ["after a quote inside a quoted field", 'a,"b"']
  ]
  const nexts: ReadonlyArray<[string, string | undefined]> = [
    ["a quote", '"'],
    ["the delimiter", ","],
    ["CR", "\r"],
    ["LF", "\n"],
    ["another character", "x"],
    ["end of input", undefined]
  ]

  for (const strict of [false, true]) {
    for (const [state, prefix] of prefixes) {
      for (const [kind, next] of nexts) {
        it(`${strict ? "strict" : "lenient"}: ${state}, then ${kind} in the next piece`, () => {
          const pieces = next === undefined ? [prefix] : [prefix, next, "z\n"]
          const whole = feedAll([pieces.join("")], { delimiter: ",", strict })
          assert.deepStrictEqual(feedAll(pieces, { delimiter: ",", strict }), whole)
          // every other split point gives the same result too
          const text = pieces.join("")
          for (let cut = 0; cut <= text.length; cut++) {
            assert.deepStrictEqual(feedAll([text.slice(0, cut), text.slice(cut)], { delimiter: ",", strict }), whole)
          }
        })
      }
    }
  }

  it("CR and LF split across pieces are one line break", () => {
    const expected = { rows: [["a"], ["b"]], error: undefined }
    assert.deepStrictEqual(feedAll(["a\r", "\nb\r\n"], { delimiter: "," }), expected)
    assert.deepStrictEqual(feedAll(["a\r", "\n", "b\r", "\n"], { delimiter: "," }), expected)
    // two CRs are two line breaks, with a blank row between them
    assert.deepStrictEqual(feedAll(["a\r", "\r", "\nb"], { delimiter: "," }), { rows: [["a"], [], ["b"]], error: undefined })
  })

  it("a quote right after a CR line break opens a quoted field", () => {
    for (const pieces of [['a\r"b,c"\r'], ["a\r", '"b,c"\r'], ["a\r", '"', 'b,c"\r']]) {
      assert.deepStrictEqual(feedAll(pieces, { delimiter: ",", strict: true }), { rows: [["a"], ["b,c"]], error: undefined })
    }
  })

  it("a closing quote at end of input ends the field and the record", () => {
    for (const strict of [false, true]) {
      assert.deepStrictEqual(feedAll(['"a"'], { delimiter: ",", strict }), { rows: [["a"]], error: undefined })
      assert.deepStrictEqual(feedAll(['x,"a', '"'], { delimiter: ",", strict }), { rows: [["x", "a"]], error: undefined })
    }
  })

  it("an unterminated quoted field at end of input keeps its text (lenient) or fails (strict)", () => {
    assert.deepStrictEqual(feedAll(["a\n", '"b'], { delimiter: "," }), { rows: [["a"], ["b"]], error: undefined })
    assert.deepStrictEqual(feedAll(["a\n", '"b'], { delimiter: ",", strict: true }), { rows: [["a"]], error: "MalformedFile" })
  })

  it("a line break at end of input adds no row; a delimiter at end of input ends a record", () => {
    for (const end of ["\n", "\r", "\r\n"]) {
      assert.deepStrictEqual(feedAll(["a" + end], { delimiter: "," }), { rows: [["a"]], error: undefined })
    }
    assert.deepStrictEqual(feedAll(["a", ","], { delimiter: "," }), { rows: [["a"]], error: undefined })
    assert.deepStrictEqual(feedAll(["a\n", ","], { delimiter: "," }), { rows: [["a"], []], error: undefined })
  })
})

describe("csv parser: lifecycle", () => {
  it("a sep= line split across pieces sets the delimiter and is not a row", () => {
    const p = parser()
    assert.deepStrictEqual(p.feed("se", false), [])
    assert.deepStrictEqual(p.feed("p=;\r", false), [])
    assert.deepStrictEqual(aoa(rowsOf(p.feed("\na;b,c\n", true))), [["a", "b,c"]])
  })

  it("error offsets count the sep= line", () => {
    // `x` follows the closing quote at index 9: 6 chars of `sep=;\n`, then `"a"`
    for (const pieces of [['sep=;\n"a"x'], ["sep=;\n", '"a"x'], ["sep", '=;\n"a"', "x"]]) {
      const p = parser({ strict: true })
      const results = pieces.map((piece, k) => p.feed(piece, k === pieces.length - 1))
      assert.strictEqual(malformedOffset(results[results.length - 1]!), 9)
    }
    // an unterminated quote is reported at the end of input: 6 + 2 chars
    assert.strictEqual(malformedOffset(parser({ strict: true }).feed('sep=;\n"a', true)), 8)
    // without a sep= line the offsets start at the first character
    assert.strictEqual(malformedOffset(parser({ strict: true }).feed('"a"x', true)), 3)
  })

  it("a sep= line whose CR and LF are split, before a sample of more than 16 Ki chars", () => {
    const body = "a;b\n".repeat(SNIFF_CHARS / 4)
    const p = parser()
    assert.deepStrictEqual(p.feed("sep=|\r", false), [])
    const rows = rowsOf(p.feed("\n" + body, false))
    assert.strictEqual(rows.length, SNIFF_CHARS / 4)
    assert.deepStrictEqual(aoa(rows)[0], ["a;b"])
  })

  it("with fewer than 16 Ki chars, sniffing waits for the final piece", () => {
    const p = parser()
    assert.deepStrictEqual(p.feed("a;b\n", false), [])
    assert.deepStrictEqual(p.feed("c;d", false), [])
    assert.deepStrictEqual(aoa(rowsOf(p.feed("", true))), [["a", "b"], ["c", "d"]])
  })

  it("sniffing starts once 16 Ki chars are collected", () => {
    const text = "a;b\n".repeat(SNIFF_CHARS / 4)
    const p = parser()
    assert.deepStrictEqual(p.feed(text.slice(0, -1), false), [])
    const rows = rowsOf(p.feed(text.slice(-1), false))
    assert.strictEqual(rows.length, SNIFF_CHARS / 4)
    assert.deepStrictEqual(aoa(rows)[0], ["a", "b"])
  })

  it("after an error, every feed returns that same error", () => {
    const p = parser({ delimiter: ",", strict: true })
    assert.deepStrictEqual(aoa(rowsOf(p.feed('a\n"b"x\nc\n', false))), [["a"]])
    const error = p.error
    assert.strictEqual(error?._tag, "MalformedFile")
    assert.strictEqual(p.feed("d\n", false), error)
    assert.strictEqual(p.feed("e\n", true), error)
    assert.strictEqual(p.error, error)
  })

  it("an error with no rows before it is returned at once", () => {
    const p = parser({ delimiter: ",", strict: true })
    const error = p.feed('"b"x', false)
    assert.isTrue(isError(error))
    assert.strictEqual(p.error, error)
    assert.strictEqual(p.feed("", true), error)
  })

  it("after the final piece, later feeds give no rows", () => {
    const p = parser({ delimiter: "," })
    assert.deepStrictEqual(aoa(rowsOf(p.feed("a\nb", true))), [["a"], ["b"]])
    assert.deepStrictEqual(p.feed("c\n", false), [])
    assert.deepStrictEqual(p.feed("d\n", true), [])
    assert.strictEqual(p.error, undefined)
  })
})

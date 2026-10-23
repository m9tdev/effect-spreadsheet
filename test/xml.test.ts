import { readdirSync, readFileSync, statSync } from "node:fs"
import { Effect, Exit, Stream } from "effect"
import fc from "fast-check"
import { describe, expect, it } from "vitest"
import { open } from "../src/internal/zip/archive.ts"
import { parse, XmlDecoder } from "../src/internal/xml/parse.ts"
import { type XmlHandler, XmlTokenizer } from "../src/internal/xml/tokenizer.ts"
import { isSpreadsheetError, type SpreadsheetError } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { failureOf } from "./helpers/run.ts"

const limits = SpreadsheetLimits.defaults

type Event = readonly ["open", string, ReadonlyArray<readonly [string, string]>] | readonly ["close", string] | readonly ["text", string]

/** Tokenizes `pieces` and returns the events (consecutive text joined) or the error, plus the work count. */
const tokenize = (
  pieces: ReadonlyArray<string>,
  l: SpreadsheetLimits.Limits = limits
): { events: Array<Event>; error: SpreadsheetError | undefined; work: number; total: number } => {
  const events: Array<Event> = []
  const handler: XmlHandler = {
    open: (name, attributes) => {
      events.push(["open", name, attributes])
    },
    close: (name) => {
      events.push(["close", name])
    },
    text: (value) => {
      const last = events[events.length - 1]
      if (last !== undefined && last[0] === "text") events[events.length - 1] = ["text", last[1] + value]
      else events.push(["text", value])
    }
  }
  const tokenizer = new XmlTokenizer(handler, l)
  let error: SpreadsheetError | undefined
  for (const piece of pieces) {
    error = tokenizer.feed(piece)
    if (error !== undefined) break
  }
  error ??= tokenizer.end()
  return { events, error, work: tokenizer.work, total: pieces.reduce((n, p) => n + p.length, 0) }
}

const one = (doc: string, l?: SpreadsheetLimits.Limits) => tokenize([doc], l)

const splitAt = (s: string, cuts: ReadonlyArray<number>): Array<string> => {
  const points = [...new Set(cuts.map((c) => Math.max(0, Math.min(s.length, c))))].sort((a, b) => a - b)
  const out: Array<string> = []
  let last = 0
  for (const p of points) {
    out.push(s.slice(last, p))
    last = p
  }
  out.push(s.slice(last))
  return out
}

const SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!-- a comment with <tags> & stuff -->
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref='A1:B2'/>
  <sheetData>
    <row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t xml:space="preserve"> a &amp; b &lt;&gt; &quot;q&quot; &apos;s&apos; &#233;&#x1F600; </t></is></c></row>
    <x:row r="2" hidden="1"><x:c r="A2"><x:f>SUM(A1:A1)</x:f><x:v>1.5</x:v></x:c><x:c r="B2" t="str"><x:v><![CDATA[<raw> & text]]></x:v></x:c></x:row>
  </sheetData>
  <?some processing instruction?>
</worksheet>
`

describe("xml tokenizer: documents", () => {
  it("tokenizes a worksheet-like document", () => {
    const { events, error } = one(SHEET)
    expect(error).toBeUndefined()
    const tags = events.filter((e) => e[0] !== "text" || e[1].trim() !== "")
    expect(tags).toEqual([
      ["open", "worksheet", [
        ["xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"],
        ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"]
      ]],
      ["open", "dimension", [["ref", "A1:B2"]]],
      ["close", "dimension"],
      ["open", "sheetData", []],
      ["open", "row", [["r", "1"], ["spans", "1:2"]]],
      ["open", "c", [["r", "A1"], ["t", "s"]]],
      ["open", "v", []],
      ["text", "0"],
      ["close", "v"],
      ["close", "c"],
      ["open", "c", [["r", "B1"], ["t", "inlineStr"]]],
      ["open", "is", []],
      ["open", "t", [["xml:space", "preserve"]]],
      ["text", ` a & b <> "q" 's' é😀 `],
      ["close", "t"],
      ["close", "is"],
      ["close", "c"],
      ["close", "row"],
      ["open", "row", [["r", "2"], ["hidden", "1"]]],
      ["open", "c", [["r", "A2"]]],
      ["open", "f", []],
      ["text", "SUM(A1:A1)"],
      ["close", "f"],
      ["open", "v", []],
      ["text", "1.5"],
      ["close", "v"],
      ["close", "c"],
      ["open", "c", [["r", "B2"], ["t", "str"]]],
      ["open", "v", []],
      ["text", "<raw> & text"],
      ["close", "v"],
      ["close", "c"],
      ["close", "row"],
      ["close", "sheetData"],
      ["close", "worksheet"]
    ])
  })

  it("'>' inside attribute values doesn't end the tag", () => {
    const { events, error } = one(`<a x="1>2" y='<b>'/>`)
    expect(error).toBeUndefined()
    expect(events).toEqual([["open", "a", [["x", "1>2"], ["y", "<b>"]]], ["close", "a"]])
  })

  it("any splitting of the input gives the same events, with work linear in the input", () => {
    fc.assert(
      fc.property(fc.array(fc.nat(SHEET.length), { maxLength: 40 }), (cuts) => {
        const whole = one(SHEET)
        const split = tokenize(splitAt(SHEET, cuts))
        expect(split.error).toBeUndefined()
        expect(split.events).toEqual(whole.events)
        expect(split.work).toBeLessThanOrEqual(6 * SHEET.length)
      }),
      { numRuns: 300 }
    )
    // one character at a time
    const single = tokenize([...SHEET])
    expect(single.events).toEqual(one(SHEET).events)
    expect(single.work).toBeLessThanOrEqual(6 * SHEET.length)
  })

  it("a long text, comment and CDATA fed one character at a time stay linear", () => {
    const doc = `<a>${"x".repeat(50_000)}<!--${"-x".repeat(20_000)}--><![CDATA[${"]".repeat(20_000)}]]><b c="${"q".repeat(30_000)}"/></a>`
    const result = tokenize([...doc])
    expect(result.error).toBeUndefined()
    expect(result.work).toBeLessThanOrEqual(6 * doc.length)
    expect(result.events.map((e) => e[0])).toEqual(["open", "text", "open", "close", "close"])
  })
})

/** The parts of an error that must not depend on how the input was split. */
const errorKey = (error: SpreadsheetError | undefined) =>
  error === undefined
    ? undefined
    : error._tag === "MalformedFile"
    ? [error._tag, error.reason]
    : error._tag === "LimitExceeded"
    ? [error._tag, error.limit]
    : [error._tag]

const XMLISH = fc.constantFrom(
  "<", ">", "/", "a", "b", " ", "=", '"', "'", "&", ";", "#", "x", "!", "-", "[", "?", "]", "\r", "\n", "\t",
  "amp", "<a>", "</a>", "<![CDATA[", "-->", "<!--", "<?", "?>", "]]>", "é", "&#10;", "&#13;", "&lt;"
)

describe("xml tokenizer: split invariance and cost", () => {
  it("a comment opener at the end of a piece can't close the comment (<!--> stays open)", () => {
    const doc = "<a><!-->x--></a>"
    const whole = one(doc)
    expect(whole.error).toBeUndefined()
    expect(whole.events).toEqual([["open", "a", []], ["close", "a"]])
    for (let cut = 0; cut <= doc.length; cut++) {
      const split = tokenize([doc.slice(0, cut), doc.slice(cut)])
      expect([split.events, errorKey(split.error)], `cut at ${cut}`).toEqual([whole.events, errorKey(whole.error)])
    }
  })

  it("random XML-ish input gives the same events and error however it is split", () => {
    fc.assert(
      fc.property(fc.array(XMLISH, { maxLength: 120 }), fc.array(fc.nat(800), { maxLength: 12 }), (parts, cuts) => {
        const doc = parts.join("")
        const whole = one(doc)
        const split = tokenize(splitAt(doc, cuts))
        expect([split.events, errorKey(split.error)]).toEqual([whole.events, errorKey(whole.error)])
      }),
      { numRuns: 3000 }
    )
  })

  it("many comments, PIs and CDATA sections in one piece cost linear work", () => {
    const doc = "<a>" + "<!-- c --><?p i?><![CDATA[d]]>".repeat(20_000) + "</a>"
    const { error, work, events } = one(doc)
    expect(error).toBeUndefined()
    expect(events.filter((e) => e[0] === "text").length).toBe(1)
    expect(work).toBeLessThanOrEqual(6 * doc.length)
  })

})

describe("xml tokenizer: rejected input", () => {
  const malformed: Record<string, ReadonlyArray<string>> = {
    "a DTD": [`<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><x>&a;</x>`],
    "a DTD split across pieces": [`<x/><!DOC`, `TYPE x>`],
    "billion laughs": [`<!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;">]><lolz>&lol2;</lolz>`],
    "an unknown entity": [`<a>&nbsp;</a>`],
    "an entity without ';'": [`<a>&amp</a>`],
    "an invalid character reference": [`<a>&#0;</a>`],
    "a surrogate character reference": [`<a>&#xD800;</a>`],
    "a reference past U+10FFFF": [`<a>&#x110000;</a>`],
    "mismatched tags": [`<a><b></a></b>`],
    "an end tag without a start": [`<a/></b>`],
    "a truncated document": [`<a><b>text`],
    "a truncated tag": [`<a><b x="1`],
    "an empty document": [``],
    "only a prolog": [`<?xml version="1.0"?>`],
    "text outside the root": [`junk<a/>`],
    "two root elements": [`<a/><b/>`],
    "a duplicate attribute": [`<a x="1" x="2"/>`],
    "an unquoted attribute": [`<a x=1/>`],
    "an attribute without a value": [`<a x/>`],
    "a bad element name": [`<1a/>`],
    "an unsupported declaration": [`<!ELEMENT a ANY><a/>`],
    "CDATA outside the root": [`<![CDATA[x]]><a/>`]
  }
  for (const [name, pieces] of Object.entries(malformed)) {
    it(`${name} is MalformedFile`, () => {
      const { error } = tokenize(pieces)
      expect(error?._tag, JSON.stringify(error)).toBe("MalformedFile")
    })
  }

  it("deep nesting stops at maxXmlDepth after a bounded amount of work", () => {
    const doc = "<a>".repeat(100_000) + "</a>".repeat(100_000)
    const { error, work } = one(doc)
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxXmlDepth", max: 128 })
    expect(work).toBeLessThan(2_000)
  })

  it("an attribute flood stops at maxXmlAttributes", () => {
    const attrs = Array.from({ length: 1000 }, (_, i) => `a${i}="${i}"`).join(" ")
    const { error } = one(`<x ${attrs}/>`)
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxXmlAttributes", max: 512 })
  })

  it("long text, attribute values, tags and comments stop at maxStringLength with bounded work", () => {
    const small = { ...limits, maxStringLength: 1000 }
    const cases = [
      `<a>${"x".repeat(1_000_000)}</a>`,
      `<a b="${"x".repeat(5_000)}"/>`,
      `<a b="${"x".repeat(1_000_000)}`,
      `<a><!--${"x".repeat(1_000_000)}`,
      `<a><![CDATA[${"x".repeat(1_000_000)}`
    ]
    for (const doc of cases) {
      // fed in 4 KiB pieces, like inflated zip output
      const pieces = Array.from({ length: Math.ceil(doc.length / 4096) }, (_, k) => doc.slice(k * 4096, (k + 1) * 4096))
      const { error, work } = tokenize(pieces, small)
      expect(error, doc.slice(0, 20)).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
      expect(work, doc.slice(0, 20)).toBeLessThan(200_000)
    }
  })

  it("random XML-ish input never throws, ends in events or a typed error, with linear work", () => {
    const alphabet = fc.constantFrom("<", ">", "/", "a", "b", " ", "=", '"', "'", "&", ";", "#", "x", "!", "-", "[", "?", "]", "amp", "<a>", "</a>", "<![CDATA[", "-->", "é")
    fc.assert(
      fc.property(fc.array(alphabet, { maxLength: 200 }), fc.array(fc.nat(400), { maxLength: 10 }), (parts, cuts) => {
        const doc = parts.join("")
        const result = tokenize(splitAt(doc, cuts))
        if (result.error !== undefined) expect(isSpreadsheetError(result.error)).toBe(true)
        expect(result.work).toBeLessThanOrEqual(6 * doc.length + 20)
      }),
      { numRuns: 2000 }
    )
  })
})

describe("xml bytes", () => {
  const bytes = (s: string) => new TextEncoder().encode(s)
  const run = (chunks: ReadonlyArray<Uint8Array>) => {
    const events: Array<string> = []
    const handler: XmlHandler = {
      open: (name) => void events.push(`<${name}>`),
      close: (name) => void events.push(`</${name}>`),
      text: (value) => void events.push(value)
    }
    return Effect.runPromise(Effect.exit(parse(Stream.fromIterable(chunks), handler, limits, "part.xml"))).then((exit) => ({ exit, events }))
  }

  it("decodes UTF-8 split anywhere and UTF-16 with a byte-order mark", async () => {
    const doc = `<a>é😀z</a>`
    const utf8 = bytes(doc)
    for (let cut = 0; cut <= utf8.length; cut++) {
      const { exit, events } = await run([utf8.subarray(0, cut), utf8.subarray(cut)])
      expect(Exit.isSuccess(exit)).toBe(true)
      expect(events.join("")).toBe("<a>é😀z</a>")
    }
    const le = new Uint8Array([0xFF, 0xFE, ...[...doc].flatMap((ch) => {
      const units: Array<number> = []
      for (let i = 0; i < ch.length; i++) units.push(ch.charCodeAt(i) & 0xFF, ch.charCodeAt(i) >> 8)
      return units
    })])
    const { exit, events } = await run([le.subarray(0, 1), le.subarray(1, 5), le.subarray(5)])
    expect(Exit.isSuccess(exit)).toBe(true)
    expect(events.join("")).toBe("<a>é😀z</a>")
  })

  it("invalid or truncated UTF-8 is MalformedFile", async () => {
    for (const chunks of [[new Uint8Array([0x3C, 0x61, 0x3E, 0xFF, 0x3C, 0x2F, 0x61, 0x3E])], [bytes("<a>"), new Uint8Array([0xC3])]]) {
      const { exit } = await run(chunks)
      expect(failureOf(exit)).toMatchObject({ _tag: "MalformedFile", part: "part.xml" })
    }
    const decoder = new XmlDecoder()
    expect(decoder.push(new Uint8Array([0x3C]))).toBe("")
    expect(decoder.end()).toBe("<")
  })

  it("parses every XML part of the committed corpus workbooks, hidden-row fixtures included, in linear work", async () => {
    let parts = 0
    for (const dir of ["test/corpus/generated", "test/corpus/hidden-rows", "test/corpus/hidden-rows/renamed"]) {
      for (const name of readdirSync(dir)) {
        if (!statSync(`${dir}/${name}`).isFile()) continue
        const bytes = new Uint8Array(readFileSync(`${dir}/${name}`))
        if (!(bytes[0] === 0x50 && bytes[1] === 0x4B)) continue
        const archive = open(bytes, { ...limits, maxCompressionRatio: 10_000 })
        if ("_tag" in archive) throw new Error(`${name}: ${archive._tag}`)
        for (const entry of archive.entries.filter((e) => /\.(xml|rels)$/.test(e.name))) {
          const handler: XmlHandler = { open: () => undefined, close: () => undefined, text: () => undefined }
          const exit = await Effect.runPromise(Effect.exit(parse(archive.read(entry), handler, limits, entry.name)))
          expect(Exit.isSuccess(exit), `${name}: ${entry.name}: ${JSON.stringify(failureOf(exit))}`).toBe(true)
          if (Exit.isSuccess(exit)) expect(exit.value.work).toBeLessThanOrEqual(6 * entry.uncompressedSize + 64)
          parts++
        }
      }
    }
    expect(parts).toBeGreaterThan(50)
  })
})

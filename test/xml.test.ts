import { readdirSync, readFileSync, statSync } from "node:fs"
import { Effect, Exit, Stream } from "effect"
import * as Result from "effect/Result"
import fc from "fast-check"
import { describe, expect, it } from "vitest"
import { open } from "../src/internal/zip/archive.ts"
import { parse, xmlText } from "../src/internal/xml/parse.ts"
import { isName, tagBound, type XmlHandler, XmlTokenizer } from "../src/internal/xml/tokenizer.ts"
import { isSpreadsheetError, MalformedFile, type SpreadsheetError } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { counter, metered } from "./helpers/meter.ts"
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
  const characters = counter("characters")
  const tokenizer = new XmlTokenizer(handler, l, undefined, characters)
  let error: SpreadsheetError | undefined
  for (const piece of pieces) {
    error = tokenizer.feed(piece)
    if (error !== undefined) break
  }
  error ??= tokenizer.end()
  return { events, error, work: characters.total(), total: pieces.reduce((n, p) => n + p.length, 0) }
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
    ? [error._tag, error.reason, error.offset]
    : error._tag === "LimitExceeded"
    ? [error._tag, error.limit]
    : [error._tag]

const XMLISH = fc.constantFrom(
  "<", ">", "/", "a", "b", " ", "=", '"', "'", "&", ";", "#", "x", "!", "-", "[", "?", "]", "\r", "\n", "\t",
  "amp", "<a>", "</a>", "<![CDATA[", "-->", "<!--", "<?", "?>", "]]>", "é", "&#10;", "&#13;", "&lt;", "<!D", "OCTYPE"
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

  it("a declaration that starts like <!DOCTYPE gives the same error wherever the input is split", () => {
    const docs = [
      "<root><!D</root>",
      "<root><!DOC x></root>",
      "<root><!DOCTYPE x></root>",
      "<!DOCTYPE x><a/>",
      "<a/><!DOCTYPEx>",
      "<a><!Dx/>y</a>",
      "<a><!DOCTYPE",
      "<a><!DO"
    ]
    for (const doc of docs) {
      const whole = one(doc)
      expect(whole.error, doc).toMatchObject({ _tag: "MalformedFile" })
      for (let cut = 0; cut <= doc.length; cut++) {
        const split = tokenize([doc.slice(0, cut), doc.slice(cut)])
        expect([split.events, errorKey(split.error)], `${doc} cut at ${cut}`).toEqual([whole.events, errorKey(whole.error)])
      }
      const perChar = tokenize([...doc])
      expect([perChar.events, errorKey(perChar.error)], `${doc} one character per piece`).toEqual([
        whole.events,
        errorKey(whole.error)
      ])
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

  it("errors carry the offset of the tag or text they are about", () => {
    expect(one("<a></b>").error).toMatchObject({ _tag: "MalformedFile", offset: 3 })
    expect(one("<a/>  junk").error).toMatchObject({ _tag: "MalformedFile", offset: 4 })
    expect(one("<a>x &bad; y</a>").error).toMatchObject({ _tag: "MalformedFile", offset: 3 })
    expect(one('<a><b c="&bad;"/></a>').error).toMatchObject({ _tag: "MalformedFile", offset: 3 })
    expect(one("<a/><!DOCTYPE x>").error).toMatchObject({ _tag: "MalformedFile", offset: 4 })
    expect(tokenize(["<a/><!DOC", "TYPE x>"]).error).toMatchObject({ _tag: "MalformedFile", offset: 4 })
  })

  it("many comments, PIs and CDATA sections in one piece cost linear work", () => {
    const doc = "<a>" + "<!-- c --><?p i?><![CDATA[d]]>".repeat(20_000) + "</a>"
    const { error, work, events } = one(doc)
    expect(error).toBeUndefined()
    expect(events.filter((e) => e[0] === "text").length).toBe(1)
    expect(work).toBeLessThanOrEqual(6 * doc.length)
  })

})

describe("xml tokenizer: markup continued across pieces", () => {
  /** Checks that `pieces` give the same events and error as the joined document in one piece. */
  const sameAsWhole = (pieces: ReadonlyArray<string>) => {
    const split = tokenize(pieces)
    const whole = one(pieces.join(""))
    expect([split.events, errorKey(split.error)], JSON.stringify(pieces)).toEqual([whole.events, errorKey(whole.error)])
    return split
  }

  it("the construct after '<' can stay undecided for several pieces", () => {
    const comment = sameAsWhole(["<r>", "<", "!", "-", "-x--><a/>", "</r>"])
    expect(comment.error).toBeUndefined()
    expect(comment.events).toEqual([["open", "r", []], ["open", "a", []], ["close", "a"], ["close", "r"]])
    const cdata = sameAsWhole(["<r>", "<!", "[CDA", "TA[x]]><a/>", "</r>"])
    expect(cdata.error).toBeUndefined()
    expect(cdata.events).toEqual([["open", "r", []], ["text", "x"], ["open", "a", []], ["close", "a"], ["close", "r"]])
  })

  it("input that ends inside any kind of markup is truncated, at the end offset", () => {
    for (const doc of ["<a/><", "<a/><!-", "<a", "<a b='x", "<a></a", "<a/><?p", "<a/><!-- x", "<a><![CDATA[x]]"]) {
      for (const pieces of [[doc], [...doc]]) {
        expect(tokenize(pieces).error, JSON.stringify(pieces)).toMatchObject({
          _tag: "MalformedFile",
          reason: "truncated inside a tag",
          offset: doc.length
        })
      }
    }
  })

  it("a start tag's open quote carries across pieces; end tags ignore quotes", () => {
    const quoted = sameAsWhole(['<a b="x>', 'y">', "</a>"])
    expect(quoted.error).toBeUndefined()
    expect(quoted.events).toEqual([["open", "a", [["b", "x>y"]]], ["close", "a"]])
    const end = sameAsWhole(["<a>", "</a", '">'])
    expect(end.error).toMatchObject({ _tag: "MalformedFile", reason: `end tag </a"> does not match <a>`, offset: 3 })
  })

  it("a terminator split across three pieces ends the comment, PI or CDATA", () => {
    const cases: ReadonlyArray<readonly [ReadonlyArray<string>, ReadonlyArray<Event>]> = [
      [["<a><!--", "-", "-", ">", "</a>"], [["open", "a", []], ["close", "a"]]],
      [["<a><!-- x -", "-", ">", "</a>"], [["open", "a", []], ["close", "a"]]],
      [["<a><?p ", "?", ">", "</a>"], [["open", "a", []], ["close", "a"]]],
      [["<a><![CDATA[x]", "]", ">", "</a>"], [["open", "a", []], ["text", "x"], ["close", "a"]]]
    ]
    for (const [pieces, events] of cases) {
      const result = sameAsWhole(pieces)
      expect(result.error, JSON.stringify(pieces)).toBeUndefined()
      expect(result.events, JSON.stringify(pieces)).toEqual(events)
    }
  })
})

describe("xml tokenizer: one root element", () => {
  it("a second root after a self-closing or closed root is rejected", () => {
    for (const doc of ["<a/><b/>", "<a></a><b/>", "<a/><a></a>"]) {
      const whole = one(doc)
      expect(whole.error, doc).toMatchObject({ _tag: "MalformedFile", reason: "more than one root element" })
      for (let cut = 0; cut <= doc.length; cut++) {
        expect(errorKey(tokenize([doc.slice(0, cut), doc.slice(cut)]).error), `${doc} cut at ${cut}`).toEqual(errorKey(whole.error))
      }
    }
  })

  it("a document with only a declaration, a PI and a comment has no root element", () => {
    const doc = `<?xml version="1.0"?>\n<?p i?>\n<!-- c -->\n`
    expect(one(doc).error).toMatchObject({ _tag: "MalformedFile", reason: "no root element", offset: doc.length })
  })

  it("whitespace, comments and PIs after the root are accepted", () => {
    const { error, events } = one("<a/> \n\t<!-- c --><?p i?>\r\n")
    expect(error).toBeUndefined()
    expect(events).toEqual([["open", "a", []], ["close", "a"]])
  })
})

describe("xml tokenizer: element text and skipped subtrees", () => {
  type Directed = readonly ["open", string] | readonly ["close", string, string | undefined] | readonly ["text", string]

  /**
   * Tokenizes `pieces`, asking for each element what `choose` says: skip
   * it, collect its text (with a prior length, if a number), or nothing.
   */
  const directed = (
    pieces: ReadonlyArray<string>,
    choose: (name: string) => "skip" | "collect" | "omit" | number | undefined,
    l: SpreadsheetLimits.Limits = limits
  ) => {
    const events: Array<Directed> = []
    const handler: XmlHandler = {
      open: (name, _, element) => {
        events.push(["open", name])
        const choice = choose(name)
        if (choice === "skip") element.skip()
        else if (choice === "collect") element.collect()
        else if (choice === "omit") element.omitText()
        else if (typeof choice === "number") element.collect(choice)
      },
      close: (name, text) => {
        events.push(["close", name, text])
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
    return { events, error }
  }

  const byName = (choices: Record<string, "skip" | "collect" | "omit" | number>) => (name: string) => choices[name]

  it("a collected element's end tag brings its text: nested elements' and CDATA included, decoded and normalized, however split", () => {
    const doc = `<a><t>x &amp;<b>y</b><![CDATA[<z>]]>\r\nw</t><t/><u>q</u></a>`
    const choose = byName({ t: "collect" })
    const whole = directed([doc], choose)
    expect(whole.error).toBeUndefined()
    expect(whole.events).toEqual([
      ["open", "a"],
      ["open", "t"],
      ["text", "x &"],
      ["open", "b"],
      ["text", "y"],
      ["close", "b", undefined],
      ["text", "<z>\nw"],
      ["close", "t", "x &y<z>\nw"],
      ["open", "t"],
      ["close", "t", ""],
      // only the element it was asked for: the next one's text isn't collected
      ["open", "u"],
      ["text", "q"],
      ["close", "u", undefined],
      ["close", "a", undefined]
    ])
    for (let cut = 0; cut <= doc.length; cut++) {
      expect(directed([doc.slice(0, cut), doc.slice(cut)], choose).events, `cut at ${cut}`).toEqual(whole.events)
    }
    expect(directed([...doc], choose).events).toEqual(whole.events)
  })

  it("a collection asked for inside a collected element joins it", () => {
    const { events, error } = directed([`<t>a<t>b</t>c</t>`], byName({ t: "collect" }))
    expect(error).toBeUndefined()
    expect(events.filter((e) => e[0] === "close")).toEqual([["close", "t", undefined], ["close", "t", "abc"]])
  })

  it("a skipped subtree gives no events, not even its end tag, and its text isn't collected; what follows it comes", () => {
    const doc = `<a><t>x<s>y<t>z</t><s>v</s></s>w</t><s/><b>q</b></a>`
    const choose = byName({ t: "collect", s: "skip" })
    const whole = directed([doc], choose)
    expect(whole.error).toBeUndefined()
    expect(whole.events).toEqual([
      ["open", "a"],
      ["open", "t"],
      ["text", "x"],
      ["open", "s"],
      ["text", "w"],
      ["close", "t", "xw"],
      ["open", "s"],
      ["open", "b"],
      ["text", "q"],
      ["close", "b", undefined],
      ["close", "a", undefined]
    ])
    for (let cut = 0; cut <= doc.length; cut++) {
      expect(directed([doc.slice(0, cut), doc.slice(cut)], choose).events, `cut at ${cut}`).toEqual(whole.events)
    }
  })

  it("an element's text can be left out of collected text, its events still coming", () => {
    const doc = `<a><t>x<o>y<t>z</t><o>v</o>u</o>w</t><o><t>q</t></o><t>p</t></a>`
    const choose = byName({ t: "collect", o: "omit" })
    const whole = directed([doc], choose)
    expect(whole.error).toBeUndefined()
    expect(whole.events).toEqual([
      ["open", "a"],
      ["open", "t"],
      ["text", "x"],
      ["open", "o"],
      ["text", "y"],
      ["open", "t"],
      ["text", "z"],
      ["close", "t", undefined],
      // one inside it doesn't end it early
      ["open", "o"],
      ["text", "v"],
      ["close", "o", undefined],
      ["text", "u"],
      ["close", "o", undefined],
      ["text", "w"],
      // the text after it is collected again
      ["close", "t", "xw"],
      // a collection asked for inside it doesn't start
      ["open", "o"],
      ["open", "t"],
      ["text", "q"],
      ["close", "t", undefined],
      ["close", "o", undefined],
      // and one after it does
      ["open", "t"],
      ["text", "p"],
      ["close", "t", "p"],
      ["close", "a", undefined]
    ])
    for (let cut = 0; cut <= doc.length; cut++) {
      expect(directed([doc.slice(0, cut), doc.slice(cut)], choose).events, `cut at ${cut}`).toEqual(whole.events)
    }
    // left-out text doesn't count against the limit
    const small = { ...limits, maxStringLength: 10 }
    expect(directed([`<a><t>12345<o>${"x".repeat(10)}</o>67890</t></a>`], choose, small).error).toBeUndefined()
  })

  it("skipped content is checked as usual: the same input fails the same way", () => {
    const docs = [
      `<a><s><b></c></s></a>`,
      `<a><s>&bogus;</s></a>`,
      `<a><s><b x="1" x="2"/></s></a>`,
      `<a><s><b><c><d/></c></b></s></a>`,
      `<a><s>${"x".repeat(50)}</s></a>`,
      `<a><s><![CDATA[${"x".repeat(50)}]]></s></a>`,
      `<a><s><b>`
    ]
    const small = { ...limits, maxXmlDepth: 3, maxStringLength: 20 }
    for (const doc of docs) {
      const plain = directed([doc], () => undefined, small)
      const skipped = directed([doc], byName({ s: "skip" }), small)
      expect(plain.error, doc).toBeDefined()
      expect(skipped.error, doc).toEqual(plain.error)
    }
  })

  it("collected text over maxStringLength, the prior length counted, stops with the total", () => {
    const small = { ...limits, maxStringLength: 10 }
    // eight characters in three runs: text, CDATA, text after a nested element
    const doc = `<a><t>123<![CDATA[45]]><b/>678</t></a>`
    expect(directed([doc], byName({ t: "collect" }), small).error).toBeUndefined()
    expect(directed([doc], byName({ t: 2 }), small).error).toBeUndefined()
    expect(directed([doc], byName({ t: 3 }), small).error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", max: 10, actual: 11 })
    // the check is on the runs as they come, so text after it doesn't count
    const cdata = `<a><t>123<![CDATA[4567890]]>xyz</t></a>`
    expect(directed([cdata], byName({ t: 1 }), small).error).toMatchObject({ _tag: "LimitExceeded", actual: 11 })
    // each collection starts from its own prior
    expect(directed([`<a><t>1234567890</t><t>1234567890</t></a>`], byName({ t: "collect" }), small).error).toBeUndefined()
    // a handler without a text callback gets the text at the end tag
    const texts: Array<string | undefined> = []
    const quiet = new XmlTokenizer({ open: (_, __, element) => element.collect(), close: (_, text) => void texts.push(text) }, limits)
    expect(quiet.feed(`<a>x<b>y</b>z</a>`) ?? quiet.end()).toBeUndefined()
    // (the inner element joined the outer one)
    expect(texts).toEqual([undefined, "xyz"])
  })
})

describe("xml tokenizer: after an error", () => {
  const stop = new MalformedFile({ format: "xml", reason: "handler stop" })

  /** Handler calls with consecutive text calls joined (a split may cut a text run). */
  const joinText = (calls: ReadonlyArray<string>): Array<string> => {
    const out: Array<string> = []
    for (const call of calls) {
      const last = out[out.length - 1]
      if (last !== undefined && last.startsWith("text ") && call.startsWith("text ")) out[out.length - 1] = last + call.slice(5)
      else out.push(call)
    }
    return out
  }

  /**
   * Feeds `pieces`, ends, then feeds more and ends again. Records every
   * handler call (the handler fails on the call `stopAt` names) and every
   * result.
   */
  const run = (pieces: ReadonlyArray<string>, stopAt: string | undefined, l: SpreadsheetLimits.Limits) => {
    const calls: Array<string> = []
    const call = (event: string) => {
      calls.push(event)
      return event === stopAt ? stop : undefined
    }
    const handler: XmlHandler = {
      open: (name) => call(`open ${name}`),
      close: (name) => call(`close ${name}`),
      text: (value) => call(`text ${value}`)
    }
    const tokenizer = new XmlTokenizer(handler, l)
    const results: Array<SpreadsheetError | undefined> = pieces.map((piece) => tokenizer.feed(piece))
    results.push(tokenizer.end())
    const settled = calls.length
    results.push(tokenizer.feed("<a>more</a>"), tokenizer.end(), tokenizer.end())
    return { calls, settled, results }
  }

  const cases: ReadonlyArray<{
    readonly name: string
    readonly doc: string
    readonly stopAt?: string
    readonly limits?: SpreadsheetLimits.Limits
    readonly error: Partial<{ _tag: string; reason: string; limit: string }>
  }> = [
    { name: "a bad entity in text", doc: "<r>a<s/>b &bogus; c</r>", error: { _tag: "MalformedFile" } },
    { name: "a bad attribute", doc: "<r><s a=1/><t/></r>", error: { _tag: "MalformedFile" } },
    {
      name: "text over maxStringLength",
      doc: `<r><s/>${"x".repeat(40)}<t/></r>`,
      limits: { ...limits, maxStringLength: 10 },
      error: { _tag: "LimitExceeded", limit: "maxStringLength" }
    },
    {
      name: "a comment over maxStringLength",
      doc: `<r><s/><!--${"x".repeat(70 * 1024)}--><t/></r>`,
      limits: { ...limits, maxStringLength: 10 },
      error: { _tag: "LimitExceeded", limit: "maxStringLength" }
    },
    { name: "the open handler of a self-closing tag", doc: "<r><s/><t/></r>", stopAt: "open s", error: stop },
    { name: "the text handler", doc: "<r>a<s/>b</r>", stopAt: "text a", error: stop },
    { name: "the close handler", doc: "<r><s></s><t/></r>", stopAt: "close s", error: stop },
    { name: "the end of input inside an element", doc: "<r><s>", error: { _tag: "MalformedFile" } }
  ]

  for (const c of cases) {
    it(`${c.name}: every later feed and end returns the same error, and no handler runs after it`, () => {
      const l = c.limits ?? limits
      const whole = run([c.doc], c.stopAt, l)
      const at = whole.results.findIndex((r) => r !== undefined)
      const error = whole.results[at]
      expect(error).toMatchObject(c.error)
      for (const later of whole.results.slice(at)) expect(later).toBe(error)
      expect(whole.calls.length).toBe(whole.settled)
      if (c.stopAt !== undefined) expect(whole.calls[whole.calls.length - 1]).toBe(c.stopAt)
      // the same however the input is split, with no handler call after the error
      // (every cut of a short document, about 500 cuts of a long one)
      const step = Math.max(1, Math.floor(c.doc.length / 500))
      for (let cut = 0; cut <= c.doc.length; cut += step) {
        const split = run([c.doc.slice(0, cut), c.doc.slice(cut)], c.stopAt, l)
        const first = split.results.findIndex((r) => r !== undefined)
        const splitError = split.results[first]
        expect(splitError).toMatchObject(c.error)
        // (a limit error's `actual` is how far the scan got, which depends on the split)
        if (splitError?._tag === "MalformedFile") expect(splitError.message).toBe(error?.message)
        for (const later of split.results.slice(first)) expect(later).toBe(splitError)
        expect(split.calls.length).toBe(split.settled)
        expect(joinText(split.calls)).toEqual(joinText(whole.calls))
      }
    })
  }
})

describe("xml names", () => {
  it("accepts and rejects exactly what the original name regex did", () => {
    const NAME = /^[^\s<>/=!?"'&\-.\d][^\s<>/=!?"'&]*$/
    const chars = [..."ab:_-.09<>/=!?\"'&", "\t", "\n", "\v", "\f", "\r", " ", "\u0001", "\u001F", "\u0085", "\u00A0", "\u1680", "\u2000", "\u200A", "\u200B", "\u2028", "\u2029", "\u202F", "\u205F", "\u3000", "\uFEFF", "é"]
    fc.assert(
      fc.property(fc.array(fc.constantFrom(...chars), { maxLength: 6 }), (parts) => {
        const name = parts.join("")
        expect(isName(name), JSON.stringify(name)).toBe(NAME.test(name))
      }),
      { numRuns: 5000 }
    )
  })
})

describe("xml tokenizer: normalization (XML 1.0)", () => {
  it("line ends in text and CDATA become LF (§2.11), however split; &#13; stays", () => {
    const doc = "<a>x\r\ny\rz\n<![CDATA[p\r\nq\rr]]>&#13;</a>"
    const expected = [["open", "a", []], ["text", "x\ny\nz\np\nq\nr\r"], ["close", "a"]]
    expect(one(doc).events).toEqual(expected)
    for (let cut = 0; cut <= doc.length; cut++) expect(tokenize([doc.slice(0, cut), doc.slice(cut)]).events).toEqual(expected)
  })

  it("literal tab, LF and CR in attribute values become spaces (§3.3.3); references stay", () => {
    const { events } = one('<a v="1\t2\n3\r\n4\r5" w="&#9;&#10;&#13;"/>')
    expect(events[0]).toEqual(["open", "a", [["v", "1 2 3 4 5"], ["w", "\t\n\r"]]])
  })

  it("numeric references may have any number of leading zeros", () => {
    expect(one("<a>&#00000065;&#x0000000042;&#0;9</a>").error).toMatchObject({ _tag: "MalformedFile" })
    expect(one("<a>&#00000065;&#x0000000042;</a>").events[1]).toEqual(["text", "AB"])
    expect(one("<a>&#00000000;</a>").error).toMatchObject({ _tag: "MalformedFile" })
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

  it("a tag complete in one piece with markup longer than maxStringLength stops", () => {
    const small = { ...limits, maxStringLength: 10 }
    const { error } = tokenize([`<a${" ".repeat(200_000)}/>`], small)
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", actual: 200_002 })
  })

  it("a duplicate among many attributes is found (past the linear check, in the Set)", () => {
    const attrs = Array.from({ length: 20 }, (_, i) => `a${i}="${i}"`).join(" ")
    expect(one(`<x ${attrs} a17="again"/>`).error).toMatchObject({ _tag: "MalformedFile", reason: "duplicate attribute a17 in <x>" })
    expect(one(`<x ${attrs} a3="again"/>`).error).toMatchObject({ _tag: "MalformedFile", reason: "duplicate attribute a3 in <x>" })
    expect(one(`<x ${attrs}/>`).error).toBeUndefined()
  })

  it("CDATA longer than maxStringLength stops, with its content's length", () => {
    const small = { ...limits, maxStringLength: 1000 }
    const { error } = one(`<a><![CDATA[${"x".repeat(5000)}]]></a>`, small)
    expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", actual: 5000 })
  })

  it("text longer than maxStringLength stops before the handler sees it, in one piece or across pieces", () => {
    const small = { ...limits, maxStringLength: 10 }
    const whole = one(`<a>${"x".repeat(50)}</a>`, small)
    expect(whole.error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", actual: 50 })
    expect(whole.events).toEqual([["open", "a", []]])
    const split = tokenize([`<a>${"y".repeat(8)}`, `${"y".repeat(8)}<b/></a>`], small)
    expect(split.error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", actual: 16 })
    expect(split.events).toEqual([["open", "a", []]])
    // a text of exactly maxStringLength is still delivered
    expect(one(`<a>${"z".repeat(10)}</a>`, small)).toMatchObject({ error: undefined, events: [["open", "a", []], ["text", "z".repeat(10)], ["close", "a"]] })
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

  describe("each part of a tag, and each comment, PI and CDATA section, is bounded on its own", () => {
    const small = { ...limits, maxStringLength: 10 }
    const over = (actual: number) => ({ _tag: "LimitExceeded", limit: "maxStringLength", max: 10, actual })
    /** Every split into two pieces, and one character at a time. */
    const splits = (doc: string): Array<Array<string>> => [
      ...Array.from({ length: doc.length + 1 }, (_, k) => [doc.slice(0, k), doc.slice(k)]),
      [...doc]
    ]

    it("an attribute value up to maxStringLength is read, one longer stops with its length, however split", () => {
      const fits = `<a b="${"x".repeat(10)}" c='${"y".repeat(10)}'/>`
      for (const pieces of splits(fits)) {
        expect(tokenize(pieces, small).events).toEqual([["open", "a", [["b", "x".repeat(10)], ["c", "y".repeat(10)]]], ["close", "a"]])
      }
      const long = `<r><s/><a b="1" c="${"x".repeat(11)}"/></r>`
      expect(one(long, small).error).toMatchObject(over(11))
      for (const pieces of splits(long)) {
        const { error, events } = tokenize(pieces, small)
        expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
        expect(events).toEqual([["open", "r", []], ["open", "s", []], ["close", "s"]])
      }
      // checked where the piece ends, before the rest of a long value arrives
      expect(tokenize([`<a b="${"x".repeat(15)}`, `${"x".repeat(1000)}"/>`], small).error).toMatchObject(over(15))
    })

    it("the markup before, between and after the values (names, '=', whitespace) is bounded like a value", () => {
      // "a" and nine spaces and "b=": 12
      expect(one(`<a${" ".repeat(9)}b="1"/>`, small).error).toMatchObject(over(12))
      expect(one(`<a${" ".repeat(6)}b="1"/>`, small).error).toBeUndefined()
      // between two values: " c" and nine spaces and "=": 12
      expect(one(`<a b="1" c${" ".repeat(9)}="2"/>`, small).error).toMatchObject(over(12))
      // after the last value: eleven spaces and "/"
      expect(one(`<a b="1"${" ".repeat(11)}/>`, small).error).toMatchObject(over(12))
      expect(one(`<a b="1"${" ".repeat(9)}/>`, small).error).toBeUndefined()
      // an element name, in a start or an end tag: the '/' of an end or
      // self-closing tag counts too
      expect(one(`<${"e".repeat(9)}></${"e".repeat(9)}>`, small).error).toBeUndefined()
      expect(one(`<${"e".repeat(9)}/>`, small).error).toBeUndefined()
      expect(one(`<${"e".repeat(10)}>`, small).error).toMatchObject({ _tag: "MalformedFile" })
      expect(one(`<${"e".repeat(10)}/>`, small).error).toMatchObject(over(11))
      expect(one(`<a></${"e".repeat(10)}>`, small).error).toMatchObject(over(11))
      expect(one(`<a></a${" ".repeat(10)}>`, small).error).toMatchObject(over(12))
      expect(one(`<a></a${" ".repeat(8)}>`, small).error).toBeUndefined()
      for (const pieces of splits(`<a b="1"${" ".repeat(11)}/>`)) {
        expect(tokenize(pieces, small).error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
      }
    })

    it("a part carried across pieces counts for its own tag only", () => {
      // nine characters of a value end the first piece; the next tag's first part is ten
      const doc = `<r><a b="${"x".repeat(9)}"/><${"d".repeat(7)} e="1"/></r>`
      const cut = doc.indexOf("x") + 9
      expect(tokenize([doc.slice(0, cut), doc.slice(cut)], small).error).toBeUndefined()
    })

    it("a tag past maxStringLength in all is read when each part is within its bound", () => {
      // 5 000 attributes, about 55 000 characters, with a maxStringLength of 10
      const attributes = Array.from({ length: 5_000 }, (_, i) => `a${i}="${i % 1000}"`)
      const doc = `<a ${attributes.join(" ")}/>`
      expect(doc.length).toBeGreaterThan(50_000)
      expect(doc.length - 2).toBeLessThanOrEqual(tagBound(10))
      const many = { ...small, maxXmlAttributes: 5_000 }
      for (const pieces of [[doc], splitAt(doc, [7, 30_000, 30_001, 49_999])]) {
        const { error, events } = tokenize(pieces, many)
        expect(error).toBeUndefined()
        expect(events[0]![0] === "open" && events[0]![2].length).toBe(5_000)
      }
    })

    describe("a whole tag is bounded by maxStringLength + 64 KiB, or twice maxStringLength if more", () => {
      it("the bound", () => {
        expect(tagBound(40)).toBe(40 + 65_536)
        expect(tagBound(65_536)).toBe(131_072)
        expect(tagBound(262_144)).toBe(524_288)
      })

      /** `<a .../>` with `length` characters between '<' and '>', of values and markup within `max` each. */
      const tagOf = (length: number, max: number): string => {
        const value = "v".repeat(Math.max(1, max - 20))
        let body = "a"
        // progress: each attribute adds more than one character, until the rest is at most max
        for (let k = 0; length - 1 - body.length > max; k++) body += ` a${k}="${value}"`
        const rest = length - 1 - body.length
        body += rest >= 6 ? ` z="${"w".repeat(rest - 5)}"` : " ".repeat(rest)
        return `<${body}/>`
      }

      for (const max of [40, 262_144]) {
        it(`a tag of exactly the bound is read, one more character fails, whole or in pieces (maxStringLength ${max})`, () => {
          const l = { ...limits, maxStringLength: max, maxXmlAttributes: 10_000 }
          const bound = tagBound(max)
          const fits = tagOf(bound, max)
          const over = tagOf(bound + 1, max)
          expect([fits.length, over.length]).toEqual([bound + 2, bound + 3])
          const cuts = [[], [1], [bound - 7], [bound], [bound + 1], [bound + 2], [bound / 2, bound - 3]]
          for (const cut of [...cuts, Array.from({ length: Math.ceil(bound / 4096) }, (_, k) => (k + 1) * 4096)]) {
            const read = tokenize(splitAt(fits, cut), l)
            expect(read.error, String(cut)).toBeUndefined()
            expect(read.events.length).toBe(2)
            const failed = tokenize(splitAt(over, cut), l)
            expect(failed.error, String(cut)).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", max })
            expect(failed.events).toEqual([])
          }
          // in one piece, with its length
          expect(one(over, l).error).toMatchObject({ actual: bound + 1 })
        })
      }

      it("at a value's opening quote, the tag up to it may reach the bound but not pass it", () => {
        const l = { ...limits, maxStringLength: 40, maxXmlAttributes: 10_000 }
        const bound = tagBound(40)
        // `a` and values up to `bound` characters before a last value's quote
        const upTo = (length: number) => {
          const head = tagOf(length - 2, 40).slice(1, -2)
          return `<${head} z="v"/>`
        }
        // `bound` characters before the quote: checked again at the '>', 4 characters later
        expect(one(upTo(bound), l).error).toMatchObject({ _tag: "LimitExceeded", actual: bound + 4 })
        // one more: stops at the quote
        expect(one(upTo(bound + 1), l).error).toMatchObject({ _tag: "LimitExceeded", actual: bound + 1 })
      })

      it("a tag's length counts for that tag only, after one that ran over pieces", () => {
        const l = { ...limits, maxStringLength: 40, maxXmlAttributes: 10_000 }
        const first = tagOf(60_000, 40)
        const second = tagOf(10_000, 40)
        const doc = `<r>${first}${second}</r>`
        // the first tag has 59 000 of its characters in the first piece
        expect(tokenize([doc.slice(0, 59_004), doc.slice(59_004)], l).error).toBeUndefined()
      })

      it("a tag over the bound fails on it before the value count it reaches later, however split", () => {
        // values of 20 characters: about 2 340 values up to the bound, which comes before the 2 401st
        const values = Array.from({ length: 3_000 }, (_, i) => `a${i}="${"v".repeat(20)}"`).join(" ")
        const doc = `<a ${values}/>`
        const l = { ...limits, maxStringLength: 40, maxXmlAttributes: 2_400 }
        for (const cut of [[], [70_000], [1, 30_000, 65_000]]) {
          expect(tokenize(splitAt(doc, cut), l).error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
        }
        // in one piece, its length up to the quote where it is checked
        const error = one(doc, l).error
        expect(error).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
        expect(error?._tag === "LimitExceeded" && error.actual).toBeGreaterThan(tagBound(40))
        expect(error?._tag === "LimitExceeded" && error.actual).toBeLessThan(tagBound(40) + 30)
        // a count reached before the bound still fails on the count
        expect(one(doc, { ...l, maxXmlAttributes: 2_000 }).error).toMatchObject({ _tag: "LimitExceeded", limit: "maxXmlAttributes" })
      })

      it("a root tag with namespace declarations is read with a low maxStringLength", () => {
        const root = `<worksheet xmlns="urn:a:main" xmlns:r="urn:a:rel" xmlns:mc="urn:a:mc" xmlns:x14ac="urn:a:x14ac" ` +
          `xmlns:xr="urn:a:xr" mc:Ignorable="x14ac xr" xr:uid="{00000000-0001-0000-0000-000000000000}">`
        expect(root.length).toBeGreaterThan(2 * 40)
        const doc = `${root}<sheetData/></worksheet>`
        for (const pieces of [[doc], splitAt(doc, [50]), [...doc]]) {
          expect(tokenize(pieces, { ...limits, maxStringLength: 40 }).error).toBeUndefined()
        }
      })
    })

    it("the values of a tag are counted against maxXmlAttributes as they are scanned", () => {
      const three = { ...limits, maxXmlAttributes: 3 }
      expect(one(`<a b="1" c='2' d="3"/>`, three).error).toBeUndefined()
      const four = `<a b="1" b="2" c="3" d="4"/>`
      // the count stops it before the duplicate is seen, however it is split
      for (const pieces of splits(four)) {
        expect(tokenize(pieces, three).error).toMatchObject({ _tag: "LimitExceeded", limit: "maxXmlAttributes", max: 3, actual: 4 })
      }
      // quotes count only in start tags
      expect(one(`<a></a>`, { ...limits, maxXmlAttributes: 0 }).error).toBeUndefined()
      expect(one(`<a b="1"/>`, { ...limits, maxXmlAttributes: 0 }).error).toMatchObject({ _tag: "LimitExceeded", limit: "maxXmlAttributes", actual: 1 })
    })

    it("an over-long part stops the tag before other errors in it", () => {
      expect(one(`<a b="1" b="${"x".repeat(11)}"/>`, small).error).toMatchObject(over(11))
      expect(one(`<1a b="${"x".repeat(11)}"/>`, small).error).toMatchObject(over(11))
      expect(one(`<a b="&bogus;" c="${"x".repeat(11)}"/>`, small).error).toMatchObject(over(11))
    })

    it("a comment's, PI's or CDATA section's content up to maxStringLength is read, one longer stops", () => {
      const docs = (n: number) => [`<a><!--${"c".repeat(n)}--></a>`, `<?${"p".repeat(n)}?><a/>`, `<a><![CDATA[${"d".repeat(n)}]]></a>`]
      for (const doc of docs(10)) {
        for (const pieces of splits(doc)) expect(tokenize(pieces, small).error, JSON.stringify(pieces)).toBeUndefined()
      }
      for (const doc of docs(11)) {
        expect(one(doc, small).error).toMatchObject(over(11))
        for (const pieces of splits(doc)) {
          expect(tokenize(pieces, small).error, JSON.stringify(pieces)).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength" })
        }
      }
      // the start of a terminator at the end of a piece is not counted yet
      expect(tokenize([`<a><![CDATA[${"d".repeat(10)}]]`, `></a>`], small).error).toBeUndefined()
      expect(tokenize([`<a><![CDATA[${"d".repeat(11)}]`, `]></a>`], small).error).toMatchObject(over(11))
      // unterminated, at the end of a piece: only what is surely content counts (all
      // but the last characters, which could start the terminator), so one past
      // the bound fails and the bound itself is truncated input
      const five = { ...limits, maxStringLength: 5 }
      const unterminated: Array<[string, number]> = [["<r><!--", 2], ["<r><?", 1], ["<r><![CDATA[", 2]]
      for (const [opener, allowance] of unterminated) {
        for (const at of [0, 2, 3 + opener.length]) {
          const fits = opener + "x".repeat(5 + allowance)
          const past = opener + "x".repeat(6 + allowance)
          expect(tokenize([fits.slice(0, at), fits.slice(at)], five).error, fits).toMatchObject({ _tag: "MalformedFile" })
          expect(tokenize([past.slice(0, at), past.slice(at)], five).error, past).toMatchObject({ _tag: "LimitExceeded", limit: "maxStringLength", actual: 6 })
        }
      }
      // CDATA outside the root that is too long stops on its length
      expect(one(`<![CDATA[${"d".repeat(11)}]]><a/>`, small).error).toMatchObject(over(11))
    })
  })

  it("a tag's characters are examined twice (scan and parse) and its '>' once, plus one per extra piece the scan runs over", () => {
    // '<' (1), then the 8 characters of `a b="1"/` and '>' (9), then the 8 again
    expect(one(`<a b="1"/>`).work).toBe(18)
    expect(tokenize([`<a b="`, `1"/>`]).work).toBe(19)
    expect(tokenize([`<`, `a b="`, `1"/>`]).work).toBe(19)
  })

  it("an attribute without a value is found in a tag read in place, whatever follows it in the piece", () => {
    expect(one(`<r><a x/><b y="1"/></r>`).error).toMatchObject({ _tag: "MalformedFile", reason: "attribute without a value in <a>" })
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

  it("a trailing incomplete UTF-8 sequence or odd UTF-16 byte is MalformedFile (the final flush catches it)", async () => {
    const utf8 = await run([bytes("<a/>"), new Uint8Array([0xC3])])
    expect(failureOf(utf8.exit)).toMatchObject({ _tag: "MalformedFile", reason: "invalid text encoding", offset: 4 })
    const le = new Uint8Array([0xFF, 0xFE, 0x3C, 0x00, 0x61, 0x00, 0x2F, 0x00, 0x3E, 0x00, 0x20])
    const utf16 = await run([le])
    expect(failureOf(utf16.exit)).toMatchObject({ _tag: "MalformedFile", reason: "invalid text encoding" })
    // the platform checks UTF-16 and doesn't say where
    expect(failureOf(utf16.exit)).not.toHaveProperty("offset")
    // the same documents without the extra byte are fine
    expect(Exit.isSuccess((await run([bytes("<a/>")])).exit)).toBe(true)
    expect(Exit.isSuccess((await run([le.subarray(0, 10)])).exit)).toBe(true)
  })

  it("invalid or truncated UTF-8 is MalformedFile, at the byte offset of the first invalid sequence in the part", async () => {
    const cases: Array<[Array<Uint8Array>, number]> = [
      [[new Uint8Array([0x3C, 0x61, 0x3E, 0xFF, 0x3C, 0x2F, 0x61, 0x3E])], 3],
      [[bytes("<a>"), new Uint8Array([0xC3])], 3],
      // a sequence split across pieces, then cut short; after a byte-order mark
      [[new Uint8Array([0xEF, 0xBB, 0xBF]), bytes("<a>é"), new Uint8Array([0xE2, 0x82]), bytes("</a>")], 8]
    ]
    for (const [chunks, offset] of cases) {
      const { exit } = await run(chunks)
      expect(failureOf(exit)).toMatchObject({ _tag: "MalformedFile", part: "part.xml", offset })
      expect(failureOf(exit)?.message).toBe(`Malformed xml file in part.xml at offset ${offset}: invalid text encoding`)
    }
    const decoder = xmlText(undefined)
    expect(decoder.push(new Uint8Array([0x3C]))).toEqual(Result.succeed(""))
    expect(decoder.end()).toEqual(Result.succeed("<"))
  })

  it("parses every XML part of the committed corpus workbooks, hidden-row fixtures included, in linear work", async () => {
    let parts = 0
    for (const dir of ["test/corpus/generated", "test/corpus/hidden-rows", "test/corpus/hidden-rows/renamed"]) {
      for (const name of readdirSync(dir)) {
        if (!statSync(`${dir}/${name}`).isFile()) continue
        const bytes = new Uint8Array(readFileSync(`${dir}/${name}`))
        if (!(bytes[0] === 0x50 && bytes[1] === 0x4B)) continue
        const archive = Result.getOrThrow(open(bytes, limits))
        for (const entry of archive.entries.filter((e) => /\.(xml|rels)$/.test(e.name))) {
          const handler: XmlHandler = { open: () => undefined, close: () => undefined, text: () => undefined }
          const characters = counter("characters")
          const exit = await Effect.runPromise(Effect.exit(metered(characters, parse(archive.read(entry, "test"), handler, limits, entry.name))))
          expect(Exit.isSuccess(exit), `${name}: ${entry.name}: ${JSON.stringify(failureOf(exit))}`).toBe(true)
          expect(characters.total()).toBeLessThanOrEqual(6 * entry.uncompressedSize + 64)
          parts++
        }
      }
    }
    expect(parts).toBeGreaterThan(50)
  })
})

describe("xml decoder: encoding sniff", () => {
  const decodeAll = (pieces: ReadonlyArray<ReadonlyArray<number>>): string | SpreadsheetError => {
    const decoder = xmlText("part.xml")
    let text = ""
    for (const piece of pieces) {
      const out = decoder.push(new Uint8Array(piece))
      if (Result.isFailure(out)) return out.failure
      text += out.success
    }
    const rest = decoder.end()
    return Result.isSuccess(rest) ? text + rest.success : rest.failure
  }
  const utf16 = (s: string, littleEndian: boolean): Array<number> =>
    [...s].flatMap((ch) => {
      const units: Array<number> = []
      for (let i = 0; i < ch.length; i++) {
        const lo = ch.charCodeAt(i) & 0xFF
        const hi = ch.charCodeAt(i) >> 8
        units.push(...(littleEndian ? [lo, hi] : [hi, lo]))
      }
      return units
    })

  it("a UTF-16 byte-order mark split one byte at a time picks the encoding (LE and BE)", () => {
    for (const [mark, littleEndian] of [[[0xFF, 0xFE], true], [[0xFE, 0xFF], false]] as const) {
      const body = utf16("<a>é😀</a>", littleEndian)
      expect(decodeAll([[mark[0]], [mark[1]], body])).toBe("<a>é😀</a>")
      // the held byte joins the next piece
      expect(decodeAll([[mark[0]], [mark[1], ...body]])).toBe("<a>é😀</a>")
    }
  })

  it("a UTF-8 byte-order mark split 1|2 is dropped", () => {
    expect(decodeAll([[0xEF], [0xBB, 0xBF], [0x3C, 0x61, 0x2F, 0x3E]])).toBe("<a/>")
  })

  it("empty input decodes to an empty string, with or without empty pieces", () => {
    expect(decodeAll([])).toBe("")
    expect(decodeAll([[], []])).toBe("")
  })

  it("a single byte then the end is decoded at the end: ASCII is text, an invalid byte is MalformedFile", () => {
    expect(decodeAll([[0x3C]])).toBe("<")
    expect(decodeAll([[0xFF]])).toMatchObject({ _tag: "MalformedFile", reason: "invalid text encoding", part: "part.xml" })
    // a lone first byte of a UTF-16 mark is not a mark
    expect(decodeAll([[0xFE]])).toMatchObject({ _tag: "MalformedFile" })
  })

  it("held bytes are not decoded until the encoding is known", () => {
    const decoder = xmlText(undefined)
    expect(decoder.push(new Uint8Array([0x3C]))).toEqual(Result.succeed(""))
    expect(decoder.push(new Uint8Array([0x61]))).toEqual(Result.succeed(""))
    expect(decoder.push(new Uint8Array([0x3E]))).toEqual(Result.succeed("<a>"))
    expect(decoder.push(new Uint8Array([0x3C]))).toEqual(Result.succeed("<"))
    expect(decoder.end()).toEqual(Result.succeed(""))
  })

  it("a U+FEFF after a 4-byte first character is text when the first 3 bytes come alone", () => {
    const input = [...new TextEncoder().encode("\u{1F600}\uFEFF<a/>")]
    expect(decodeAll([input.slice(0, 3), input.slice(3)])).toBe("\u{1F600}\uFEFF<a/>")
  })

  it("a UTF-8 byte-order mark split across pieces is dropped once: a second one is text, however the input is split", () => {
    const input = [0xEF, 0xBB, 0xBF, 0xEF, 0xBB, 0xBF, 0x3C, 0x61, 0x2F, 0x3E]
    for (let cut = 0; cut <= input.length; cut++) {
      expect(decodeAll([input.slice(0, cut), input.slice(cut)])).toBe("\uFEFF<a/>")
    }
  })
})

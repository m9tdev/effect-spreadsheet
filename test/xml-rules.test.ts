import { describe, expect, it } from "vitest"
import { MAX_STEPS, parsePath, rulesHandler, type XmlRule, type XmlRules } from "../src/internal/xml/rules.ts"
import { XmlTokenizer } from "../src/internal/xml/tokenizer.ts"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"

const malformed = (reason: string) => new MalformedFile({ format: "xml", reason })

/** Tokenizes `xml` with `rules` (each `on` path logging its opens and closes); the log and the error. */
const run = (
  xml: string,
  paths: ReadonlyArray<string>,
  options: { forbid?: Record<string, string>; skip?: ReadonlyArray<string>; rule?: (path: string, log: Array<string>) => XmlRule } = {}
): { log: Array<string>; error: SpreadsheetError | undefined } => {
  const log: Array<string> = []
  const on: Record<string, XmlRule> = {}
  for (const path of paths) {
    on[path] = options.rule?.(path, log) ?? {
      open: (attributes) => {
        log.push(`open ${path} ${attributes.map(([k, v]) => `${k}=${v}`).join(",")}`.trimEnd())
      },
      close: (text) => {
        log.push(`close ${path}${text === undefined ? "" : ` "${text}"`}`)
      }
    }
  }
  const rules: XmlRules = {
    on,
    malformed,
    ...(options.forbid === undefined ? {} : { forbid: options.forbid }),
    ...(options.skip === undefined ? {} : { skip: options.skip })
  }
  const tokenizer = new XmlTokenizer(rulesHandler(rules), SpreadsheetLimits.defaults)
  const error = tokenizer.feed(xml) ?? tokenizer.end()
  return { log, error }
}

describe("xml path rules: paths", () => {
  it("splits a path into local names after / (a child) or // (any depth)", () => {
    expect(parsePath("//a/b//c")).toEqual([["a", true], ["b", false], ["c", true]])
    expect(parsePath("/root")).toEqual([["root", false]])
  })

  it("a path that doesn't start with a slash, has an empty step or a prefixed name, or is empty, is a programming error", () => {
    for (const bad of ["a", "ab", "//a/", "//a///b", "/", "//", "", "//x:a", "a//b"]) expect(() => parsePath(bad), bad).toThrow(/XML path/)
  })

  it(`the rules may have at most ${MAX_STEPS} steps in all`, () => {
    const on = (steps: number): Record<string, XmlRule> => ({ [Array.from({ length: steps }, (_, k) => `/e${k}`).join("")]: {} })
    expect(() => rulesHandler({ on: on(MAX_STEPS), malformed })).not.toThrow()
    expect(() => rulesHandler({ on: on(MAX_STEPS + 1), malformed })).toThrow(/more than 30/)
    // forbid and skip paths count too
    expect(() => rulesHandler({ on: on(MAX_STEPS), forbid: { "//x": "x" }, malformed })).toThrow(/more than 30 path steps/)
    expect(() => rulesHandler({ on: on(MAX_STEPS), skip: ["//x"], malformed })).toThrow(/more than 30 path steps/)
    // paths that start alike share their steps
    expect(() => rulesHandler({ on: { ...on(MAX_STEPS), "/e0": {}, "/e0/e1": {} }, malformed })).not.toThrow()
  })

  it(`the rules may have at most ${MAX_STEPS} paths in all, even on fewer steps`, () => {
    const paths = Array.from({ length: 10 }, (_, k) => `//p${k}`)
    const on = Object.fromEntries(paths.map((p) => [p, {}]))
    const forbid = Object.fromEntries(paths.map((p) => [p, "x"]))
    expect(() => rulesHandler({ on, forbid, skip: paths, malformed })).not.toThrow()
    expect(() => rulesHandler({ on: { ...on, "//p10": {} }, forbid, skip: paths, malformed })).toThrow(/more than 30 paths/)
  })
})

describe("xml path rules: matching", () => {
  it("/name: the root element, or a child of the step before; never deeper", () => {
    const { log } = run(`<a><b/><x><b/></x><a><b/></a></a>`, ["/a", "/a/b"])
    expect(log).toEqual(["open /a", "open /a/b", "close /a/b", "close /a"])
  })

  it("//name: any element, the root included, or a descendant at any depth of the step before", () => {
    const { log } = run(`<a><b/><x><b/></x></a>`, ["//a", "//b", "//a//b"])
    expect(log).toEqual([
      "open //a",
      "open //b",
      "open //a//b",
      "close //b",
      "close //a//b",
      "open //b",
      "open //a//b",
      "close //b",
      "close //a//b",
      "close //a"
    ])
    expect(run(`<b/>`, ["//b", "//a//b"]).log).toEqual(["open //b", "close //b"])
  })

  it("names are local names: a prefix in the document doesn't matter", () => {
    expect(run(`<x:a xmlns:x="u"><y:b/></x:a>`, ["/a/b"]).log).toEqual(["open /a/b", "close /a/b"])
  })

  it("descendant steps match below any wrapper, child steps only right under their parent", () => {
    const xml = `<r><list><w><item n="1"/></w><item n="2"/></list><item n="3"/></r>`
    expect(run(xml, ["//list//item"]).log.filter((l) => l.startsWith("open"))).toEqual(["open //list//item n=1", "open //list//item n=2"])
    expect(run(xml, ["//list/item"]).log.filter((l) => l.startsWith("open"))).toEqual(["open //list/item n=2"])
    expect(run(xml, ["/r/item"]).log.filter((l) => l.startsWith("open"))).toEqual(["open /r/item n=3"])
  })

  it("an element nested in one of its own kind matches again, and the outer one still matches after it", () => {
    const xml = `<r><l><i n="1"/><l><i n="2"/></l><i n="3"/></l></r>`
    expect(run(xml, ["//l//i"]).log.filter((l) => l.startsWith("open"))).toEqual(["open //l//i n=1", "open //l//i n=2", "open //l//i n=3"])
    // a step that matched once is still there for the deeper elements
    expect(run(`<a><a><b/></a></a>`, ["//a/b"]).log).toEqual(["open //a/b", "close //a/b"])
  })

  it("an element matching several paths: their opens in the order the paths are given, then their closes in the same order", () => {
    const { log } = run(`<a><b>t</b></a>`, ["//b", "/a/b", "//a//b"])
    expect(log).toEqual(["open //b", "open /a/b", "open //a//b", "close //b", "close /a/b", "close //a//b"])
  })

  it("siblings with the same name at the same depth match by their own parent", () => {
    const xml = `<a><m><i n="1"/><i n="2"/></m><l><i n="3"/><i n="4"/></l><m><i n="5"/></m></a>`
    expect(run(xml, ["/a/l/i"]).log.filter((l) => l.startsWith("open"))).toEqual(["open /a/l/i n=3", "open /a/l/i n=4"])
  })

  it("a rule without open or close is fine", () => {
    expect(run(`<a><b/></a>`, ["//b"], { rule: () => ({}) })).toEqual({ log: [], error: undefined })
  })
})

describe("xml path rules: forbid", () => {
  it("an element on a forbidden path is MalformedFile with that path's reason, before any open", () => {
    const { log, error } = run(`<a><b/><b><b/></b></a>`, ["//b"], { forbid: { "//x": "an x", "//b//b": "a b in a b" } })
    expect(error).toMatchObject({ _tag: "MalformedFile", reason: "a b in a b" })
    expect(log).toEqual(["open //b", "close //b", "open //b"])
  })

  it("the first forbidden path an element matches gives the reason", () => {
    const forbid = { "//a//b": "first", "//b": "second" }
    expect(run(`<a><b/></a>`, [], { forbid }).error).toMatchObject({ reason: "first" })
    expect(run(`<b/>`, [], { forbid }).error).toMatchObject({ reason: "second" })
  })

  it("an element only near a forbidden path is fine", () => {
    expect(run(`<a><c><b/></c></a>`, ["//b"], { forbid: { "/a/b": "no" } }).error).toBeUndefined()
  })
})

describe("xml path rules: what the rules ask of the tokenizer", () => {
  it("the element's collected text comes to close; a rule's error stops tokenizing", () => {
    const rule = (path: string, log: Array<string>): XmlRule =>
      path === "//t"
        ? { open: (_, element) => element.collect(), close: (text) => void log.push(`t "${text}"`) }
        : { open: () => new LimitExceeded({ limit: "maxSheets", max: 0, actual: 1 }) }
    expect(run(`<a><t>x<u>y</u>z</t></a>`, ["//t"], { rule })).toEqual({ log: [`t "xyz"`], error: undefined })
    const { log, error } = run(`<a><t>x</t><stop/><t>never</t></a>`, ["//t", "//stop"], { rule })
    expect(log).toEqual([`t "x"`])
    expect(error).toMatchObject({ _tag: "LimitExceeded" })
  })

  it("an open's error: the element's later rules don't open; a close's error stops tokenizing too", () => {
    const stop = new LimitExceeded({ limit: "maxSheets", max: 0, actual: 1 })
    const opens = run(`<a><b/></a>`, ["//b", "/a/b"], {
      rule: (path, log) => ({ open: () => (log.push(`open ${path}`), path === "//b" ? stop : undefined) })
    })
    expect(opens).toEqual({ log: ["open //b"], error: stop })
    const closes = run(`<a><b/><b/></a>`, ["//b", "/a/b"], {
      rule: (path, log) => ({ close: () => (log.push(`close ${path}`), path === "//b" ? stop : undefined) })
    })
    expect(closes).toEqual({ log: ["close //b"], error: stop })
  })

  it("a skipped element: no rules for it, inside it, nor at its end tag; the rules go on after it", () => {
    const { log, error } = run(`<a><s><b/><b><c/></b></s><b/></a>`, ["//s", "//b", "//b/c"], { skip: ["//s"], forbid: { "//s//c": "c in s" } })
    expect(error).toBeUndefined()
    expect(log).toEqual(["open //b", "close //b"])
    // the matching below and after it isn't thrown off
    expect(run(`<a><s><x/></s><x/><b><s/><x/></b></a>`, ["/a/x", "//b/x"], { skip: ["/a/s", "//b/s"] }).log)
      .toEqual(["open /a/x", "close /a/x", "open //b/x", "close //b/x"])
  })

  it("a forbidden path wins over a skip path", () => {
    expect(run(`<a><s/></a>`, [], { skip: ["//s"], forbid: { "//a/s": "no s" } }).error).toMatchObject({ reason: "no s" })
  })

  it("left-out text: the events still come, and a <t> inside asks for its text in vain", () => {
    const rule = (path: string, log: Array<string>): XmlRule =>
      path === "//o"
        ? { open: (_, element) => element.omitText() }
        : { open: (_, element) => element.collect(), close: (text) => void log.push(`t ${String(text)}`) }
    expect(run(`<a><t>1</t><o><t>2</t></o></a>`, ["//o", "//t"], { rule }).log).toEqual(["t 1", "t undefined"])
  })
})

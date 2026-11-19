import { assert, describe, it } from "@effect/vitest"
import fc from "fast-check"
import { fieldEnds, type FieldSink, lexFields, type LexState, lexState, Mode } from "../src/internal/csv/lex.ts"

type Event =
  | readonly ["field", string, string]
  | readonly ["record", string, boolean]
  | readonly ["afterQuote", number]

/**
 * Records what the lexer reports, as whole fields: the text parts of a
 * field are joined, and the character that ends it is kept. `afterQuote`
 * offsets count from the start of the whole input (pieces are fed in
 * order). `stopAt`: return false on that event (counted from 0).
 */
class Recorder implements FieldSink {
  readonly events: Array<Event> = []
  field = ""
  offset = 0
  calls = 0
  readonly stopAt: number
  readonly lenient: boolean
  constructor(stopAt = -1, lenient = true) {
    this.stopAt = stopAt
    this.lenient = lenient
  }

  private go(): boolean {
    return this.calls++ !== this.stopAt
  }

  text(s: string, from: number, to: number): boolean {
    assert.isTrue(to > from, "text parts are not empty")
    this.field += s.slice(from, to)
    return this.go()
  }

  endField(s: string, from: number, to: number): boolean {
    assert.isTrue(to >= from)
    const field = this.field + s.slice(from, to)
    this.field = ""
    this.events.push(["field", field, s[to]!])
    return this.go()
  }

  endRecord(started: boolean): boolean {
    const last = this.events[this.events.length - 1]
    assert.isTrue(last !== undefined && last[0] === "field" && (last[2] === "\n" || last[2] === "\r"))
    this.events.push(["record", last![2] as string, started])
    return this.go()
  }

  afterQuote(at: number): boolean {
    this.events.push(["afterQuote", this.offset + at])
    return this.lenient && this.go()
  }
}

/** Lexes the pieces in order with one state; returns the events and the state. */
const lex = (pieces: ReadonlyArray<string>, delimiters: ReadonlyArray<string> = [","], sink = new Recorder()) => {
  const state = lexState()
  const ends = fieldEnds(delimiters.map((d) => d.charCodeAt(0)))
  for (const piece of pieces) {
    lexFields(piece, 0, state, ends, sink)
    sink.offset += piece.length
  }
  return { events: sink.events, state, open: sink.field }
}

describe("csv lexer", () => {
  it("field ends: the delimiters, CR and LF, and nothing else", () => {
    const ends = fieldEnds([0x3B, 0x2603])
    assert.strictEqual(ends.length, 0x10000)
    const set: Array<number> = []
    ends.forEach((v, c) => {
      if (v !== 0) set.push(c)
    })
    assert.deepStrictEqual(set, [0x0A, 0x0D, 0x3B, 0x2603])
  })

  it("reports fields and records, with the character that ends each field", () => {
    const { events, state } = lex(["a,b\r\nc,\"d,\"\"e\"\"\"\n\n,\r"])
    assert.deepStrictEqual(events, [
      ["field", "a", ","],
      ["field", "b", "\r"],
      ["record", "\r", true],
      ["field", "c", ","],
      ["field", 'd,"e"', "\n"],
      ["record", "\n", true],
      // a blank line: a record with one empty field, and nothing before its line break
      ["field", "", "\n"],
      ["record", "\n", false],
      ["field", "", ","],
      ["field", "", "\r"],
      ["record", "\r", true]
    ])
    assert.deepStrictEqual(state, { mode: Mode.FieldStart, skipLF: true, recordStarted: false })
  })

  it("a record started by an empty quoted field or a delimiter has started", () => {
    assert.deepStrictEqual(lex(['""\n']).events, [["field", "", "\n"], ["record", "\n", true]])
    assert.deepStrictEqual(lex([",\n"]).events, [["field", "", ","], ["field", "", "\n"], ["record", "\n", true]])
    assert.deepStrictEqual(lex(["a\n"]).events, [["field", "a", "\n"], ["record", "\n", true]])
  })

  it("every delimiter of the set ends a field; other candidates are text", () => {
    assert.deepStrictEqual(lex(["a;b|c,d\n"], [";", "|"]).events, [
      ["field", "a", ";"],
      ["field", "b", "|"],
      ["field", "c,d", "\n"],
      ["record", "\n", true]
    ])
  })

  it("a quote opens a quoted field only at its start", () => {
    assert.deepStrictEqual(lex(['a"b,"c,d"\n']).events, [
      ["field", 'a"b', ","],
      ["field", "c,d", "\n"],
      ["record", "\n", true]
    ])
  })

  it("text after a closing quote: reported at its offset, then read leniently", () => {
    const { events } = lex(['"ab"c"d,', '"x"', "y\n"])
    assert.deepStrictEqual(events, [
      ["afterQuote", 4],
      ["field", 'abc"d', ","],
      ["afterQuote", 11],
      ["field", "xy", "\n"],
      ["record", "\n", true]
    ])
  })

  it("stops at once when the sink says so, at any event", () => {
    const text = 'a,"b""c"x\r\nd\n'
    const whole = new Recorder()
    const all = lex([text], [","], whole).events
    assert.strictEqual(whole.calls, 9)
    // progress: k increases up to the number of sink calls
    for (let k = 0; k < whole.calls; k++) {
      const sink = new Recorder(k)
      lex([text], [","], sink)
      assert.strictEqual(sink.calls, k + 1, `stopped after call ${k}`)
      // what was reported before the stop is a prefix of the full run
      assert.deepStrictEqual(sink.events, all.slice(0, sink.events.length))
    }
    // afterQuote returning false stops too
    const strict = new Recorder(-1, false)
    lex(['"a"b,c\n'], [","], strict)
    assert.deepStrictEqual(strict.events, [["afterQuote", 3]])
  })

  it("an open field and record at the end of the input are left in the state", () => {
    const { events, state, open } = lex(['a,"b\nc'])
    assert.deepStrictEqual(events, [["field", "a", ","]])
    assert.strictEqual(open, "b\nc")
    assert.deepStrictEqual(state, { mode: Mode.Quoted, skipLF: false, recordStarted: true })
    assert.deepStrictEqual(lex(["ab"]).state, { mode: Mode.Unquoted, skipLF: false, recordStarted: true })
    assert.deepStrictEqual(lex([""]).state, { mode: Mode.FieldStart, skipLF: false, recordStarted: false })
  })

  it("a CR at the end of a piece and a LF at the start of the next are one line break", () => {
    const { events } = lex(["a\r", "\nb\r", "", "\n"])
    assert.deepStrictEqual(events, [["field", "a", "\r"], ["record", "\r", true], ["field", "b", "\r"], [
      "record",
      "\r",
      true
    ]])
    // a CR then another CR is two line breaks
    assert.strictEqual(lex(["a\r", "\r"]).events.length, 4)
  })

  it("starts at `from`", () => {
    const state: LexState = lexState()
    const sink = new Recorder()
    lexFields("sep=;\na;b\n", 6, state, fieldEnds([0x3B]), sink)
    assert.deepStrictEqual(sink.events, [["field", "a", ";"], ["field", "b", "\n"], ["record", "\n", true]])
  })

  it("any split into pieces gives the same fields, records and state", () => {
    const alphabet = fc.constantFrom("a", "é", "😀", ",", ";", '"', '"', '""', "\r", "\n", "\r\n", " ")
    fc.assert(
      fc.property(
        fc.array(alphabet, { maxLength: 60 }).map((xs) => xs.join("")),
        fc.array(fc.nat(), { maxLength: 8 }),
        (text, cuts) => {
          const points = [...new Set(cuts.map((c) => c % (text.length + 1)))].sort((a, b) => a - b)
          const pieces: Array<string> = []
          let at = 0
          // progress: one cut point per iteration
          for (const p of points) {
            pieces.push(text.slice(at, p))
            at = p
          }
          pieces.push(text.slice(at))
          assert.deepStrictEqual(lex(pieces, [",", ";"]), lex([text], [",", ";"]))
        }
      ),
      { numRuns: 2000 }
    )
  })
})

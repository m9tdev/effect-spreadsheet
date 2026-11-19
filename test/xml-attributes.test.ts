import fc from "fast-check"
import { describe, expect, it } from "vitest"
import * as Attr from "../src/internal/xml/attributes.ts"
import type { XmlAttributes } from "../src/internal/xml/tokenizer.ts"

const one = (name: string, value: string): XmlAttributes => [["other", "x"], [name, value]]
const none: XmlAttributes = [["other", "x"]]

describe("typed XML attributes", () => {
  it("string: the value as written, or undefined when absent", () => {
    expect(Attr.string(one("t", "s"), "t")).toBe("s")
    expect(Attr.string(one("t", ""), "t")).toBe("")
    expect(Attr.string(one("t", "_x0041_"), "t")).toBe("_x0041_")
    expect(Attr.string(none, "t")).toBeUndefined()
    // names are matched as written
    expect(Attr.string(one("r:id", "a"), "id")).toBeUndefined()
    // the first of two (the tokenizer rejects duplicates; this is the reader's own rule)
    expect(Attr.string([["t", "a"], ["t", "b"]], "t")).toBe("a")
  })

  it("xstring: _xHHHH_ escapes decoded, one UTF-16 code unit each", () => {
    expect(Attr.xstring(one("name", "a_x0041_b"), "name")).toBe("aAb")
    expect(Attr.xstring(one("name", "_x005F_x0041_"), "name")).toBe("_x0041_")
    expect(Attr.xstring(one("name", "_xD83D__xDE00_"), "name")).toBe("\u{1F600}")
    expect(Attr.xstring(one("name", "_x00e9_ _X0041_ _x41_"), "name")).toBe("é _X0041_ _x41_")
    expect(Attr.xstring(one("name", ""), "name")).toBe("")
    expect(Attr.xstring(none, "name")).toBeUndefined()
  })

  it("bool: only 1 and true are true", () => {
    expect(Attr.bool(one("hidden", "1"), "hidden")).toBe(true)
    expect(Attr.bool(one("hidden", "true"), "hidden")).toBe(true)
    for (const value of ["0", "false", "TRUE", "True", " 1", "1 ", "01", "", "yes"]) {
      expect(Attr.bool(one("hidden", value), "hidden"), value).toBe(false)
    }
    expect(Attr.bool(none, "hidden")).toBe(false)
  })

  it("uint: ASCII digits, at most as many as max has, at most max", () => {
    const nine = 999_999_999
    expect(Attr.uint(none, "s", nine)).toBeUndefined()
    expect(Attr.uint(one("s", "0"), "s", nine)).toBe(0)
    expect(Attr.uint(one("s", "007"), "s", nine)).toBe(7)
    expect(Attr.uint(one("s", "999999999"), "s", nine)).toBe(999_999_999)
    // ten characters are more digits than max has, even with a leading zero
    expect(Attr.uint(one("s", "0999999999"), "s", nine)).toBe("bad")
    expect(Attr.uint(one("s", "1000000000"), "s", nine)).toBe("bad")
    for (const value of ["", "-1", "+1", " 1", "1 ", "1.0", "1e3", "0x10", "\u0663", "\uff11", "a", "/", ":"]) {
      expect(Attr.uint(one("s", value), "s", nine), value).toBe("bad")
    }
    // a max that isn't all nines: the value decides within the digit count
    expect(Attr.parseUint("255", 255)).toBe(255)
    expect(Attr.parseUint("256", 255)).toBe("bad")
    expect(Attr.parseUint("0255", 255)).toBe("bad")
    expect(Attr.parseUint("9", 9)).toBe(9)
    expect(Attr.parseUint("10", 9)).toBe("bad")
    expect(Attr.parseUint("0", 0)).toBe(0)
    expect(Attr.parseUint("1", 0)).toBe("bad")
    expect(Attr.parseUint("9999999999", 9_999_999_999)).toBe(9_999_999_999)
    expect(Attr.parseUint("10000000000", 9_999_999_999)).toBe("bad")
  })

  it("uint agrees with a digit-count pattern for every max of nines", () => {
    fc.assert(
      fc.property(fc.integer({ min: 1, max: 15 }), fc.oneof(fc.stringMatching(/^\d{0,17}$/), fc.string({ maxLength: 12 })), (digits, text) => {
        const max = 10 ** digits - 1
        const expected = new RegExp(`^\\d{1,${digits}}$`).test(text) ? Number(text) : "bad"
        expect(Attr.parseUint(text, max)).toBe(expected)
      })
    )
  })

  it("rel: the first prefixed attribute with that local name, whatever the prefix", () => {
    expect(Attr.rel(one("r:id", "rId1"), "id")).toBe("rId1")
    expect(Attr.rel(one("x:id", "rId2"), "id")).toBe("rId2")
    expect(Attr.rel([["x:id", "a"], ["r:id", "b"]], "id")).toBe("a")
    // unprefixed, another local name, or more than one colon: not a relationship id
    expect(Attr.rel(one("id", "rId1"), "id")).toBeUndefined()
    expect(Attr.rel(one(":id", "rId1"), "id")).toBeUndefined()
    expect(Attr.rel(one("r:Id", "rId1"), "id")).toBeUndefined()
    expect(Attr.rel(one("r:ids", "rId1"), "id")).toBeUndefined()
    expect(Attr.rel(one("a:b:id", "rId1"), "id")).toBeUndefined()
    expect(Attr.rel(none, "id")).toBeUndefined()
  })
})

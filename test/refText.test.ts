import { describe, expect, it } from "vitest"
import { needsQuotes, sheetPrefix } from "../src/internal/refText.ts"

describe("needsQuotes", () => {
  it("quotes names that read as an A1 cell reference, up to column XFD", () => {
    for (const name of ["A1", "AB1", "XFD1", "xfd1", "A0", "Ab12"]) expect(needsQuotes(name), name).toBe(true)
  })

  it("leaves names past column XFD or with more than three letters unquoted", () => {
    for (const name of ["XFE1", "xfe1", "AAAA1", "ZZZ1"]) expect(needsQuotes(name), name).toBe(false)
  })

  it("quotes R1C1-like and non-name-like names, and leaves plain names unquoted", () => {
    for (const name of ["R", "C2", "R1C1", "My Sheet", "1abc"]) expect(needsQuotes(name), name).toBe(true)
    for (const name of ["Sheet1", "Data", "_x", "Données"]) expect(needsQuotes(name), name).toBe(false)
  })

  it("decides the quoting of a sheet prefix", () => {
    expect(sheetPrefix("", ["Sheet1"])).toBe("Sheet1!")
    expect(sheetPrefix("", ["AB1"])).toBe("'AB1'!")
    expect(sheetPrefix("", ["XFE1"])).toBe("XFE1!")
  })
})

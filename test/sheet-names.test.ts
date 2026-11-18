import fc from "fast-check"
import { describe, expect, it } from "vitest"
import { uniqueSheetNames } from "../src/internal/sheetNames.ts"

const lower = (s: string) => s.toLowerCase()

describe("unique sheet names", () => {
  it("keeps the first of each name; later ones and nameless sheets get a name no other sheet has", () => {
    expect(uniqueSheetNames([])).toEqual([])
    expect(uniqueSheetNames(["A", "B"])).toEqual(["A", "B"])
    expect(uniqueSheetNames(["A", "A", "a", "A"])).toEqual(["A", "A (2)", "a (3)", "A (4)"])
    // a suffix never takes a name a later sheet has
    expect(uniqueSheetNames(["A", "A", "A (2)", "a (3)"])).toEqual(["A", "A (4)", "A (2)", "a (3)"])
    // nameless: Sheet<n> from its position on, skipping names in use
    expect(uniqueSheetNames(["A", undefined])).toEqual(["A", "Sheet2"])
    expect(uniqueSheetNames([undefined, "sheet1", undefined, "Sheet3"])).toEqual(["Sheet2", "sheet1", "Sheet4", "Sheet3"])
    // an empty name is a name: only a second one gets a suffix
    expect(uniqueSheetNames(["", ""])).toEqual(["", " (2)"])
  })

  it("many sheets with one name, or nameless sheets whose Sheet<n> are taken, cost linear time", () => {
    // a search from the start for each sheet would take seconds here; the bound is generous for a loaded machine
    const same = Array<string>(20_000).fill("A")
    const nameless = [...Array.from({ length: 10_000 }, (_, k) => `Sheet${k + 10_001}`), ...Array<undefined>(10_000).fill(undefined)]
    const start = performance.now()
    const a = uniqueSheetNames(same)
    const b = uniqueSheetNames(nameless)
    expect(performance.now() - start).toBeLessThan(2_000)
    expect(a.at(-1)).toBe("A (20000)")
    expect(b.at(-1)).toBe("Sheet30000")
    expect(new Set(b.map(lower)).size).toBe(b.length)
  })

  it("random names: every name unique in any case, first ones kept, already unique lists unchanged", () => {
    const name = fc.option(fc.constantFrom("A", "a", "A (2)", "a (3)", "Sheet1", "SHEET2", "Sheet3", "B", " (2)", ""), { nil: undefined })
    fc.assert(
      fc.property(fc.array(name, { maxLength: 12 }), (names) => {
        const out = uniqueSheetNames(names)
        expect(out.length).toBe(names.length)
        expect(new Set(out.map(lower)).size).toBe(out.length)
        const seen = new Set<string>()
        names.forEach((n, k) => {
          if (n === undefined) expect(out[k]).toMatch(/^Sheet\d+$/)
          else if (seen.has(lower(n))) expect(out[k]!.startsWith(`${n} (`)).toBe(true)
          else {
            expect(out[k]).toBe(n)
            seen.add(lower(n))
          }
        })
      }),
      { numRuns: 1000 }
    )
    fc.assert(
      fc.property(fc.uniqueArray(fc.string({ maxLength: 8 }), { selector: lower, maxLength: 12 }), (names) => {
        expect(uniqueSheetNames(names)).toEqual(names)
      }),
      { numRuns: 1000 }
    )
  })
})

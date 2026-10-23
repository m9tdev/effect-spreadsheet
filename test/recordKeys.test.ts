import { describe, it } from "@effect/vitest"
import fc from "fast-check"
import { expect } from "vitest"
import { make, makeCounted, recordKeys } from "../src/internal/recordKeys.ts"

const range = (n: number, f: (i: number) => string): Array<string> => Array.from({ length: n }, (_, i) => f(i))

/** Independent naive reference for dense headers: linear scans, suffixes tried from 1 every time. */
const reference = (names: ReadonlyArray<string | undefined>): Array<string> => {
  const used: Array<string> = []
  const take = (b: string): string => {
    let key = b
    for (let k = 1; used.includes(key); k++) key = `${b}_${k}`
    used.push(key)
    return key
  }
  const out: Array<string | undefined> = names.map((n) => (n === undefined || n === "" ? undefined : take(n)))
  return out.map((k) => k ?? take("__EMPTY"))
}

const timed = <A>(f: () => A): { readonly value: A; readonly ms: number } => {
  const start = performance.now()
  const value = f()
  return { value, ms: performance.now() - start }
}

describe("recordKeys: spec examples (dense header)", () => {
  const cases: ReadonlyArray<readonly [number, ReadonlyArray<string | undefined>, ReadonlyArray<string>]> = [
    [1, ["a", "b"], ["a", "b"]],
    [2, ["a", "a", "a"], ["a", "a_1", "a_2"]],
    [3, ["a", "a", "", "a", "", "a_1"], ["a", "a_1", "__EMPTY", "a_2", "__EMPTY_1", "a_1_1"]],
    [4, ["a", "a_1", "a"], ["a", "a_1", "a_2"]],
    [5, ["a_1", "a", "a"], ["a_1", "a", "a_2"]],
    [6, ["c", "c", "c", "c_1"], ["c", "c_1", "c_2", "c_1_1"]],
    [7, ["a", "a", "a_1_1", "a_1"], ["a", "a_1", "a_1_1", "a_1_2"]],
    [8, ["", "", ""], ["__EMPTY", "__EMPTY_1", "__EMPTY_2"]],
    [9, ["__EMPTY", ""], ["__EMPTY", "__EMPTY_1"]],
    [10, ["", "__EMPTY"], ["__EMPTY_1", "__EMPTY"]],
    [11, ["__proto__", "__proto__"], ["__proto__", "__proto___1"]],
    [12, [" a", "a", "a ", "A"], [" a", "a", "a ", "A"]],
    [13, [" ", " "], [" ", " _1"]],
    [14, ["x_0", "x", "x"], ["x_0", "x", "x_1"]],
    [15, ["x_01", "x", "x"], ["x_01", "x", "x_1"]],
    [16, ["1", "1"], ["1", "1_1"]],
    [17, ["a", ...range(99, (i) => `a_${i + 1}`), "a"], ["a", ...range(99, (i) => `a_${i + 1}`), "a_100"]],
    [18, range(16384, () => "a"), range(16384, (i) => (i === 0 ? "a" : `a_${i}`))],
    [19, range(16384, () => ""), range(16384, (i) => (i === 0 ? "__EMPTY" : `__EMPTY_${i}`))]
  ]
  for (const [n, names, keys] of cases) {
    it(`example ${n}`, () => {
      expect(recordKeys(names)).toEqual(keys)
      // the naive reference is quadratic or worse: only for the small rows
      if (names.length <= 200) expect(reference(names)).toEqual(keys)
    })
  }

  it("undefined names and holes are empty, like \"\"", () => {
    expect(recordKeys([undefined, "a", undefined])).toEqual(["__EMPTY", "a", "__EMPTY_1"])
    expect(recordKeys([, "a", ,])).toEqual(["__EMPTY", "a", "__EMPTY_1"])
    expect(recordKeys([])).toEqual([])
  })

  it("example 17 takes the ladder's free rung in one call of bounded probes", () => {
    const names = ["a", ...range(99, (i) => `a_${i + 1}`), "a"]
    const keys = makeCounted(names.map((n, i) => [i, n] as const))
    // 101 checks of the base, 100 suffix probes for the last `a` (99 fail, 1 succeeds)
    expect(keys.named[100]).toEqual([100, "a_100"])
    expect(keys.probes()).toBe(101 + 100)
  })
})

describe("recordKeys: make returns a plain RecordKeys", () => {
  it("has exactly named, headerKey and key (no probe counter)", () => {
    expect(Object.keys(make([[0, "a"]])).sort()).toEqual(["headerKey", "key", "named"])
    expect("probes" in make([])).toBe(false)
    expect(typeof makeCounted([]).probes).toBe("function")
  })
})

describe("recordKeys: lazy keys with make (spec examples 20-24)", () => {
  it("20: columns past the header get __EMPTY names", () => {
    const keys = make([[0, "a"]])
    expect([keys.key(0), keys.key(1), keys.key(2)]).toEqual(["a", "__EMPTY", "__EMPTY_1"])
    expect(keys.named).toEqual([[0, "a"]])
  })

  it("21: gap columns inside the header get no header key and are named when data appears", () => {
    const keys = make([[0, "a"], [1, undefined], [2, "c"]])
    expect(keys.named).toEqual([[0, "a"], [2, "c"]])
    expect(keys.headerKey(1)).toBeUndefined()
    expect(keys.key(0)).toBe("a") // row 1
    expect(keys.key(1)).toBe("__EMPTY") // row 2
    expect(keys.key(1)).toBe("__EMPTY") // stable on later calls
    expect(keys.headerKey(1)).toBeUndefined() // a lazy key is not a header key
    expect(keys.headerKey(2)).toBe("c")
  })

  it("22: header missing", () => {
    const keys = make([])
    expect(keys.named).toEqual([])
    expect(keys.key(1)).toBe("__EMPTY")
  })

  it("23: a far cell costs one key, not a dense row", () => {
    const keys = make([[0, "a"]])
    expect(keys.key(16383)).toBe("__EMPTY")
    expect(keys.named).toEqual([[0, "a"]])
  })

  it("24: lazy keys follow the order of first use, not column order", () => {
    const keys = make([[0, "a"]])
    expect(keys.key(2)).toBe("__EMPTY")
    expect(keys.key(1)).toBe("__EMPTY_1")
    expect(keys.key(2)).toBe("__EMPTY")
  })

  it("a sparse header: only the cells present, far columns included", () => {
    const keys = make([[3, "x"], [16383, "x"]])
    expect(keys.named).toEqual([[3, "x"], [16383, "x_1"]])
    expect(keys.headerKey(16383)).toBe("x_1")
    expect(keys.headerKey(0)).toBeUndefined()
    expect(keys.key(0)).toBe("__EMPTY")
  })

  it("header keys are allocated before lazy ones: a lazy key never takes a header name", () => {
    const keys = make([[0, ""], [1, "__EMPTY"], [2, "__EMPTY_1"]])
    expect(keys.key(0)).toBe("__EMPTY_2")
    expect(keys.key(5)).toBe("__EMPTY_3")
  })

  it("each make is a fresh run: no state leaks between runs", () => {
    const run = () => {
      const keys = make([[0, "a"], [1, "a"]])
      return [keys.key(0), keys.key(1), keys.key(2), keys.key(4)]
    }
    expect(run()).toEqual(["a", "a_1", "__EMPTY", "__EMPTY_1"])
    expect(run()).toEqual(["a", "a_1", "__EMPTY", "__EMPTY_1"])
  })

  it("violated preconditions don't throw: repeated column keeps the first, descending is put in order, bad columns ignored", () => {
    const repeated = make([[0, "a"], [0, "b"], [1, "a"]])
    expect(repeated.named).toEqual([[0, "a"], [1, "a_1"]])
    const descending = make([[2, "a"], [0, "a"], [1, "b"]])
    expect(descending.named).toEqual([[0, "a"], [1, "b"], [2, "a_1"]])
    const bad = make([[-1, "a"], [Number.NaN, "a"], [1.5, "a"], [Infinity, "a"], [0, "a"]])
    expect(bad.named).toEqual([[0, "a"]])
  })

  it("names are ordinary strings: __proto__, constructor, toString", () => {
    const keys = make([[0, "__proto__"], [1, "constructor"], [2, "toString"], [3, "__proto__"], [4, "hasOwnProperty"]])
    expect(keys.named.map(([, k]) => k)).toEqual(["__proto__", "constructor", "toString", "__proto___1", "hasOwnProperty"])
    expect(make([]).headerKey(0)).toBeUndefined()
  })
})

describe("recordKeys: hostile inputs stay linear", () => {
  // Two checks per hostile header. The work counter is exact and doesn't
  // depend on the machine: at most 3 `used` lookups per key. Timing catches
  // quadratic work the counter can't see (say, a list scan inside a lookup).
  // An absolute bound and a growth ratio (n vs 4n) both raced machine load:
  // under load linear code reached ~1.4 s and a 4x growth of 10-15x, as a
  // larger input suffers more from cache and GC pressure. So the time is
  // compared with a linear reference doing the same kind of work on the same
  // input (an entry array, a Set and two Maps of n fresh strings), measured
  // in the same process, interleaved, best of 3: load slows both alike.
  // Measured at 65 536 columns: linear code 1.0-1.8x the reference idle and
  // 0.4-3.1x under load; a quadratic that scans a list on every 4th
  // allocation 16-255x.
  const N = 65536
  const MAX_RELATIVE = 8
  const hostile: ReadonlyArray<readonly [string, (n: number) => ReadonlyArray<string>]> = [
    ["n × a", (n) => range(n, () => "a")],
    ["n × empty", (n) => range(n, () => "")],
    ["a_1 … a_(n/2-1) then n/2 × a", (n) => [...range(n / 2 - 1, (i) => `a_${i + 1}`), ...range(n / 2, () => "a")]],
    ["a, a_1 … a_(n/2-1) then n/2 × a", (n) => ["a", ...range(n / 2 - 1, (i) => `a_${i + 1}`), ...range(n / 2, () => "a")]],
    ["n × __EMPTY_k pre-filled, then n empty", (n) => [...range(n, (i) => (i === 0 ? "__EMPTY" : `__EMPTY_${i}`)), ...range(n, () => "")]]
  ]
  const allKeys = (names: ReadonlyArray<string>) => {
    const keys = makeCounted(names.map((n, i) => [i, n] as const))
    for (let i = 0; i < names.length; i++) keys.key(i)
    return keys
  }
  /** Linear work of the same shape as `allKeys`, for comparison. */
  const reference = (names: ReadonlyArray<string>): number => {
    const entries = names.map((n, i) => [i, n] as const)
    const used = new Set<string>()
    const byColumn = new Map<number, string>()
    const next = new Map<string, number>()
    // progress: one entry per iteration
    for (const [i, n] of entries) {
      const key = `${n}_${i}`
      if (!used.has(key)) used.add(key)
      byColumn.set(i, key)
      next.set(n, i)
    }
    return used.size + byColumn.size + next.size
  }
  for (const [name, header] of hostile) {
    it(`${name}, n = 16384: at most 3 lookups per key, all keys distinct`, () => {
      const names = header(16384)
      const value = allKeys(names)
      expect(value.probes()).toBeLessThanOrEqual(3 * names.length)
      const keys = names.map((_, i) => value.key(i))
      expect(new Set(keys).size).toBe(names.length)
    })

    it(`${name}, n = ${N}: at most ${MAX_RELATIVE}x the time of linear reference work`, async () => {
      const names = header(N)
      // untimed warm-up of both, so neither minimum includes JIT work
      allKeys(names)
      reference(names)
      let keysMs = Infinity
      let referenceMs = Infinity
      // progress: r increases up to 3. All 3 attempts count toward both
      // minimums, and the order alternates (reference first, then keys
      // first), so neither side always runs right after the other's garbage.
      // Runs are separated by an event-loop turn, so a slow (quadratic) run
      // fails here instead of tripping the hang watchdog. The only early exit
      // is when the ratio is far off, and only from the second attempt on, so
      // one slow first sample can't decide the result.
      for (let r = 0; r < 3; r++) {
        await new Promise((resolve) => setImmediate(resolve))
        if (r % 2 === 0) {
          referenceMs = Math.min(referenceMs, timed(() => reference(names)).ms)
          keysMs = Math.min(keysMs, timed(() => allKeys(names)).ms)
        } else {
          keysMs = Math.min(keysMs, timed(() => allKeys(names)).ms)
          referenceMs = Math.min(referenceMs, timed(() => reference(names)).ms)
        }
        if (r >= 1 && keysMs / Math.max(referenceMs, 0.5) > 4 * MAX_RELATIVE) break
      }
      expect(
        keysMs / Math.max(referenceMs, 0.5),
        `keys ${keysMs.toFixed(1)} ms, reference ${referenceMs.toFixed(1)} ms`
      ).toBeLessThan(MAX_RELATIVE)
    }, 120_000)
  }

  it("the ladder case gives the expected keys", () => {
    const names = [...range(8191, (i) => `a_${i + 1}`), ...range(8192, () => "a")]
    const keys = recordKeys(names)
    expect(keys.slice(0, 8191)).toEqual(names.slice(0, 8191))
    expect(keys[8191]).toBe("a")
    expect(keys[8192]).toBe("a_8192")
    expect(keys[keys.length - 1]).toBe(`a_${8191 + 8191}`)
  })
})

describe("recordKeys: properties", () => {
  const alphabet = fc.constantFrom<string | undefined>(
    "a", "a_1", "a_2", "a_1_1", "a_01", "_", "1", "1_1", "", undefined, " ", "__EMPTY", "__EMPTY_1", "__proto__", "__proto___1", "b"
  )
  const isSuffix = (s: string): boolean => /^[1-9][0-9]*$/.test(s)
  const derives = (key: string, base: string): boolean =>
    key === base || (key.startsWith(`${base}_`) && isSuffix(key.slice(base.length + 1)))

  it("dense headers: unique, s or s_k, first come keeps its name, matches the naive reference, deterministic, bounded", () => {
    fc.assert(
      fc.property(fc.array(alphabet, { maxLength: 60 }), (names) => {
        const counted = makeCounted(names.map((n, i) => [i, n] as const))
        const keys = names.map((_, i) => counted.key(i))
        expect(new Set(keys).size).toBe(keys.length)
        const headerSoFar = new Set<string>()
        names.forEach((name, i) => {
          if (name === undefined || name === "") {
            expect(derives(keys[i], "__EMPTY")).toBe(true)
            expect(counted.headerKey(i)).toBeUndefined()
          } else {
            expect(derives(keys[i], name)).toBe(true)
            if (!headerSoFar.has(name)) expect(keys[i]).toBe(name)
            headerSoFar.add(keys[i])
            expect(counted.headerKey(i)).toBe(keys[i])
          }
        })
        expect(keys).toEqual(reference(names))
        expect(recordKeys(names)).toEqual(keys)
        expect(recordKeys(names)).toEqual(keys)
        expect(counted.probes()).toBeLessThanOrEqual(3 * names.length)
      }),
      { numRuns: 2000 }
    )
  })

  it("sparse headers and random key calls: unique per column, stable, lazy keys in call order, bounded", () => {
    fc.assert(
      fc.property(
        fc.uniqueArray(fc.tuple(fc.nat(40), alphabet), { selector: ([c]) => c, maxLength: 30 }),
        fc.array(fc.nat(50), { maxLength: 80 }),
        (raw, calls) => {
          const header = [...raw].sort((x, y) => x[0] - y[0])
          const counted = makeCounted(header)
          const namedColumns = new Set(counted.named.map(([c]) => c))
          expect(counted.named.map(([c]) => c)).toEqual(
            header.filter(([, n]) => n !== undefined && n !== "").map(([c]) => c)
          )
          const byColumn = new Map<number, string>()
          const lazyOrder: Array<string> = []
          for (const column of calls) {
            const k = counted.key(column)
            const before = byColumn.get(column)
            if (before !== undefined) expect(k).toBe(before)
            else {
              byColumn.set(column, k)
              if (!namedColumns.has(column)) lazyOrder.push(k)
            }
            expect(counted.headerKey(column)).toBe(namedColumns.has(column) ? k : undefined)
          }
          expect(new Set(byColumn.values()).size).toBe(byColumn.size)
          // lazy keys are the smallest free __EMPTY rungs, in order of first use
          const headerKeys = new Set(counted.named.map(([, k]) => k))
          const free: Array<string> = []
          for (let k = 0; free.length < lazyOrder.length; k++) {
            const candidate = k === 0 ? "__EMPTY" : `__EMPTY_${k}`
            if (!headerKeys.has(candidate)) free.push(candidate)
          }
          expect(lazyOrder).toEqual(free)
          expect(counted.probes()).toBeLessThanOrEqual(3 * (counted.named.length + lazyOrder.length))
        }
      ),
      { numRuns: 1000 }
    )
  })
})

/**
 * Record keys for `Spreadsheet.toRecords`: one unique object key per column.
 *
 * - A header name used for the first time stays as it is.
 * - A repeated name gets the smallest free numeric suffix: `a`, `a_1`, `a_2`, ...
 * - A column without a header name (empty cell, hole in a sparse header row,
 *   column past the header) gets `__EMPTY`, `__EMPTY_1`, ..., assigned lazily
 *   on the first `key` call for that column.
 *
 * Names are compared as exact strings (no trimming, no case folding, no
 * normalization). `used` is a `Set` and the lazy table a `Map`, so names like
 * `__proto__` are ordinary. One `RecordKeys` is made per run; it keeps no
 * global state.
 *
 * Cost: every `take` is O(1) amortized (see `take`), and nothing is allocated
 * per column index, so a far cell at column 16383 costs one map entry.
 */
import { type Meter, noMeter } from "./meter.ts"

/** One header cell: its 0-based column index and its name. */
export type HeaderEntry = readonly [column: number, name: string | undefined]

export interface RecordKeys {
  /** Header-named columns and their keys, in ascending column order.
   *  These are the columns `defval` fills. */
  readonly named: ReadonlyArray<readonly [column: number, key: string]>
  /** The key of `column` if the header named it, else `undefined`. Pure lookup, O(1). */
  readonly headerKey: (column: number) => string | undefined
  /** The key for `column`: its header key if it has one; otherwise its lazy
   *  `__EMPTY…` key, assigned on the first call for that column and returned
   *  unchanged on every later call. */
  readonly key: (column: number) => string
}

const EMPTY = "__EMPTY"

const isEmptyName = (name: string | undefined): name is undefined | "" => name === undefined || name === ""

/** `header`: entries with strictly ascending, non-negative safe integer
 *  columns, one per column (the records code compacts the header row's
 *  cells first, like a reader compacts a row's). Sparse: only the cells
 *  present. Never a dense array up to the max column. The entries are
 *  trusted, not checked: a violated precondition throws nothing and every
 *  loop stays bounded, but a repeated column would be named twice. The
 *  `used` lookups of each key allocation (first checks included) go to
 *  `meter`. */
export const make = (header: Iterable<HeaderEntry>, meter: Meter = noMeter): RecordKeys => {
  const used = new Set<string>()
  /** Per base name: the next suffix to try. Invariant: `b_1 … b_{next(b)-1}` are all in `used`. */
  const next = new Map<string, number>()
  const lazy = new Map<number, string>()
  const byColumn = new Map<number, string>()

  /**
   * Allocates a key from base `b`: `b` if free, else `b_k` for the smallest
   * free `k ≥ 1`. Keys are never released, so starting at `next(b)` is exact
   * by the invariant above.
   *
   * Progress: `k` strictly increases each iteration, and the loop only
   * continues while `b_k` is already in `used`; the candidates are distinct
   * strings, so one call runs at most `used.size + 1` iterations.
   * Amortized: a key `s` can fail a suffix probe only for the one base `b`
   * with `s = b_k` (`String(k)` has no `_`), and for that base `k` only
   * grows, so every key fails at most one suffix probe per run. Each call
   * also does one check of `b` and at most one successful suffix probe, so
   * total lookups ≤ 3 × keys handed out.
   */
  const take = (b: string): string => {
    if (!used.has(b)) {
      meter.add("probes", 1)
      used.add(b)
      return b
    }
    let k = next.get(b) ?? 1
    let candidate = `${b}_${k}`
    let probes = 2
    while (used.has(candidate)) {
      k++
      candidate = `${b}_${k}`
      probes++
    }
    meter.add("probes", probes)
    next.set(b, k + 1)
    used.add(candidate)
    return candidate
  }

  // Left to right, first come first served, no look-ahead.
  // Progress: one step per entry of the (finite) header.
  const named: Array<readonly [number, string]> = []
  for (const [column, name] of header) {
    if (isEmptyName(name)) continue
    const key = take(name)
    byColumn.set(column, key)
    named.push([column, key])
  }

  const headerKey = (column: number): string | undefined => byColumn.get(column)

  const key = (column: number): string => {
    const fromHeader = byColumn.get(column)
    if (fromHeader !== undefined) return fromHeader
    const known = lazy.get(column)
    if (known !== undefined) return known
    const fresh = take(EMPTY)
    lazy.set(column, fresh)
    return fresh
  }

  return { named, headerKey, key }
}

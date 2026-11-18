import { type HeaderEntry, make } from "../../src/internal/recordKeys.ts"

/** The keys for a dense header, as if every column had data in the first
 *  data row. Equal to `make(entries).key(i)` for i = 0 … names.length - 1,
 *  called in that order. Holes count as empty. */
export const recordKeys = (names: ReadonlyArray<string | undefined>): ReadonlyArray<string> => {
  const entries: Array<HeaderEntry> = []
  // Progress: `i` increases by one per step up to `names.length`.
  for (let i = 0; i < names.length; i++) entries.push([i, names[i]])
  const keys = make(entries)
  const out: Array<string> = []
  // Progress: `i` increases by one per step up to `names.length`.
  for (let i = 0; i < names.length; i++) out.push(keys.key(i))
  return out
}

/**
 * Sheet names that are unique in a workbook, without case (Excel compares
 * sheet names that way, and so does `Workbook.findSheet`), so that every
 * sheet can be found by its name.
 */

/**
 * The names the sheets are reported under, from the names the file gives
 * (decoded, in sheet order; `undefined` for a sheet without one). The first
 * sheet with a name keeps it. A later sheet with the same name, in any case,
 * gets `Name (2)`, or the first of `Name (3)`, `Name (4)`, ... that no other
 * sheet has. A sheet without a name gets the first of `Sheet<n>`,
 * `Sheet<n+1>`, ... that no other sheet has, `n` being its 1-based position.
 * Names that are already unique come back unchanged.
 */
export const uniqueSheetNames = (names: ReadonlyArray<string | undefined>): Array<string> => {
  // every name the file gives is reserved up front, so a made-up name never
  // takes one a later sheet has
  const taken = new Set<string>()
  const keeps: Array<boolean> = []
  // progress: one name per iteration
  for (const name of names) {
    const key = name?.toLowerCase()
    const keep = key !== undefined && !taken.has(key)
    if (keep) taken.add(key)
    keeps.push(keep)
  }
  // where the search for the next `Name (n)` of a name (without case) starts,
  // and for the next `Sheet<n>`: searches only go forward, so each taken
  // name fails at most one try, and many sheets with one name cost linear time
  const nextSuffix = new Map<string, number>()
  let nextSheet = 1
  const out: Array<string> = []
  // progress: one name per iteration
  for (let k = 0; k < names.length; k++) {
    const name = names[k]
    if (keeps[k]) {
      out.push(name!)
      continue
    }
    const key = name?.toLowerCase()
    const candidate = name === undefined ? (n: number) => `Sheet${n}` : (n: number) => `${name} (${n})`
    let n = key === undefined ? Math.max(k + 1, nextSheet) : nextSuffix.get(key) ?? 2
    // progress: n increases; each failed try is a taken name
    while (taken.has(candidate(n).toLowerCase())) n++
    if (key === undefined) nextSheet = n + 1
    else nextSuffix.set(key, n + 1)
    const made = candidate(n)
    taken.add(made.toLowerCase())
    out.push(made)
  }
  return out
}

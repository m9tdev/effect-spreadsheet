# Clean-room spec: record key naming (`src/internal/recordKeys.ts`)

Status: spec for a clean-room rewrite. The implementer gets only this file plus
the repo's public tests and docs. It is written without reading the current
implementation (`src/internal/records.ts`) or any SheetJS source. Its sources are
the repo's tests (`test/records.test.ts`), the doc comment of
`Spreadsheet.toRecords` in `src/Spreadsheet.ts`, and the MIT-licensed docs and
tests listed under "Sources".

The new module must not carry a SheetJS attribution header and must not be
listed in `NOTICE`. Whoever integrates it handles deleting the old `unique()` code
and updating `NOTICE`/`LICENSE-APACHE-2.0` (other files may still need them).

## 1. Purpose

`Spreadsheet.toRecords` turns rows into plain objects keyed by a header. Each
object key must be unique, or a later column silently overwrites an earlier one
in every record. This module turns header cells into keys, one per column:

- a header name that is used for the first time stays as it is;
- a repeated header name gets a numeric suffix (`a`, `a_1`, `a_2`, …);
- a column with no header name (an empty header cell, a hole in a sparse header
  row, or a column past the header's width) gets `__EMPTY`, `__EMPTY_1`, …,
  assigned lazily, when the records code first needs a key for that column.

The module only names keys. It doesn't build records, apply `defval`, count
limits or convert cell values.

## 2. Recommendation: keep de-duplication in v1

Keep it, with the rules below. Reasons:

1. **Repeated and empty headers are common in real files.** Report exports have
   blank spacer columns, repeated `Notes`/`Total`/`Date` columns, and merged
   headers that only fill the first cell. Rejecting these files would make
   `toRecords` fail on ordinary input.
2. **Without renaming, data is lost silently.** If keys collide, the last value
   wins. csv-parse documents exactly this for its `columns` mode: "When
   multiple columns share the same name, only the last value is retained.
   Previous values are swallowed." PapaParse added automatic renaming to
   prevent it: "Duplicated field names will be automatically renamed to avoid
   values in previous fields having the same name to be overwritten."
3. **A typed error (`DuplicateHeader`) is the strict alternative, but a worse
   default.** It rejects common files (reason 1), and the only workaround it
   leaves is `header: [...]`, which means reading the header yourself first.
4. **Rows-as-arrays already exist** (`Workbook.rows`) for callers who need the
   exact cells. `toRecords` is the convenience layer, and the convenience is that
   it works on messy files without losing values.
5. **Grouping into arrays** (csv-parse `group_columns_by_name`) changes the
   value type to `Value | Array<Value>` for every key. It doesn't fit
   `Record<string, Cell.Value | D>` and it would break the typed API test.
6. **The cost is small and bounded** (section 6). The current outputs (`__EMPTY`,
   `name_1`) are pinned by tests and match the SheetJS `sheet_to_json` naming
   that the API doc promises. Users moving from SheetJS keep the same keys.

If you decide to drop it anyway: replace this module with a check that fails
with a typed `SpreadsheetError` on the first repeated non-empty header name, and
name empty/extra columns `__EMPTY_<column index>` (unique by construction, no
suffix logic needed). Four pinned tests would then have to change: "duplicate
and empty header names are made unique like SheetJS", "columns past the header
get __EMPTY names", "header row missing from sparse rows: columns get __EMPTY
names", and "defval fills only header-named columns; far cells don't widen every
record (B1 regression)" (its `__EMPTY` would become `__EMPTY_16383`). I don't
recommend this.

## 3. Interface

Headers arrive all at once: for every `header` mode the header is known before
the first data row (`"first-row"`, a row index `n`, or an explicit array).
Columns without a header name, though, get names lazily: the doc says
"A column with no header cell is named `__EMPTY_k` when a data cell first
appears in it". The test "gap columns inside the header get no defval and are
named when data appears" pins this, and a header row may be sparse (a cell at
column 16383 alone). So the module is a small stateful namer created **once per
run** (the test "works lazily on streams" re-runs a stream and expects the same
keys; state must not leak between runs):

```ts
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

/** `header`: entries with strictly ascending, non-negative integer columns.
 *  Sparse: only the cells present. Never a dense array up to the max column. */
export const make: (header: Iterable<HeaderEntry>) => RecordKeys

/** Pure helper (tests, docs): the keys for a dense header, as if every column
 *  had data in the first data row. Equal to `make(entries).key(i)` for
 *  i = 0 … names.length - 1, called in that order. */
export const recordKeys: (names: ReadonlyArray<string | undefined>) => ReadonlyArray<string>
```

How `records.ts` uses it:

- `"first-row"` / `header: n`: convert the header row's cells to entries (see
  4.6) and call `make`. If the header row is missing (sparse rows skip it), call
  `make([])`: every column gets an `__EMPTY…` key (test "header row missing from
  sparse rows").
- `header: [...]`: `make(array.map((name, i) => [i, name]))`. The API drops
  columns at or past `array.length` (test "explicit header names; extra columns
  dropped"), so the caller never calls `key` for them. Recommended: run explicit
  names through the same de-duplication, so `["x", "x"]` doesn't lose data.
  No test pins this either way.
- For each data cell that becomes a field, use `key(cell.column)`. For `defval`
  fills, iterate `named` only.

## 4. Rules

Notation: `used` is the set of keys handed out so far in this run. `b_k` is the
string `b + "_" + String(k)`, where `k` is a positive integer written in decimal
without leading zeros.

### 4.1 Allocation of one key from a base name `b`

`take(b)`:
- if `b` is not in `used`: the key is `b`;
- otherwise: the key is `b_k` for the **smallest** `k ≥ 1` with `b_k` not in
  `used`.

Add the key to `used` and return it. At most one suffix is ever appended: the
key is `b` or `b_k`, never `b_k_m` built in two steps. (A literal header
`a_1` can still become `a_1_1`, because its base is `a_1`.)

### 4.2 Header names: first come, first served, left to right

`make` processes the entries in ascending column order. An entry whose name is
non-empty gets `take(name)`. Each header gets its key before any later header
is looked at. There is no look-ahead reservation of literal names that appear
further right. This is pinned: `a,a,,a,,a_1` gives `a_1` to the second column
and `a_1_1` to the literal `a_1` (test "duplicate and empty header names are
made unique like SheetJS").

Note: PapaParse ≥ 5.4 behaves differently here. Its test "Duplicated header
names existing column" maps `c,c,c,c_1` to `c,c_2,c_3,c_1`, keeping the literal
`c_1`. We deliberately keep the pinned first-come rule, for SheetJS
compatibility and the existing test.

### 4.3 Empty names and columns without a header name

A name is **empty** if it is `undefined` or exactly `""`. Empty entries, holes
in a sparse header row, and columns past the last entry all get no header key
(`headerKey` → `undefined`, not in `named`, no `defval` fill). The first
`key(column)` call for such a column does `take("__EMPTY")` and remembers the
result for that column (in a `Map<number, string>`, not an array indexed by
column).

All header keys are allocated in `make`, before any lazy key. So a lazy key
never takes a name that a header uses: header `__EMPTY,` gives the gap column
`__EMPTY_1`.

Lazy keys are numbered in the order of the first `key` calls, which is the
order in which data cells first appear (rows in order, cells ascending within a
row). It is not column order: with header `a`, row 1 having data only at column
2 and row 2 data at column 1, column 2 is `__EMPTY` and column 1 is `__EMPTY_1`.

### 4.4 Collisions with names that already look suffixed

These need no special case beyond 4.1/4.2. `b_k` is only ever compared as a
whole string against `used`:

- a literal `a_1` that comes **before** the second `a` keeps `a_1`, and the
  second `a` then gets `a_2`;
- a literal `a_1` that comes **after** the second `a` (which took `a_1`) gets
  `a_1_1`;
- `x_0` and `x_01` are just other names: suffixes start at 1 and have no leading
  zeros, so they never collide with generated keys.

### 4.5 Strings are compared exactly

No trimming, case folding or Unicode normalization. `" a"`, `"a"`, `"a "`, `"A"`
are four distinct names. A whitespace-only name such as `" "` is **not** empty:
it stays `" "`, and a repeat becomes `" _1"`. This is the library's faithful-
values stance. Callers who want trimmed or normalized keys can map the header
themselves (PapaParse leaves this to `transformHeader`). Names such as
`__proto__`, `constructor` and `toString` are ordinary: `used` must be a `Set`
and the lazy table a `Map`, never plain-object lookups. Records are built so
that `__proto__` becomes an own key (test "__proto__ header becomes an own key,
not the prototype"). That part belongs to `records.ts`, but `recordKeys` must
not break it (PapaParse has the same case: `__proto__,__proto__,__proto__` →
`__proto__`, `__proto___1`, `__proto___2`).

### 4.6 Non-string header cells (caller's job, in `records.ts`)

`make` only takes `string | undefined`. Converting header cells is the caller's
job. Pinned by "dates in header become ISO text; values are cell values": a Date
becomes `date.toISOString()` and the number `1` becomes `"1"`. Recommended for
the rest: text as is, number `String(n)`, boolean `"true"`/`"false"`, an error
cell its code (`"#N/A"`), an Empty cell `undefined`.

### 4.7 Determinism

Keys depend only on the header entries and on the order of `key` calls. No
randomness, no global state.

## 5. Examples

`recordKeys(names)` (dense header, every column has data):

| # | header names | keys |
|---|---|---|
| 1 | `a, b` | `a, b` |
| 2 | `a, a, a` | `a, a_1, a_2` |
| 3 | `a, a, "", a, "", a_1` | `a, a_1, __EMPTY, a_2, __EMPTY_1, a_1_1` (pinned) |
| 4 | `a, a_1, a` | `a, a_1, a_2` |
| 5 | `a_1, a, a` | `a_1, a, a_2` |
| 6 | `c, c, c, c_1` | `c, c_1, c_2, c_1_1` (PapaParse: `c, c_2, c_3, c_1`) |
| 7 | `a, a, a_1_1, a_1` | `a, a_1, a_1_1, a_1_2` |
| 8 | `"", "", ""` | `__EMPTY, __EMPTY_1, __EMPTY_2` |
| 9 | `__EMPTY, ""` | `__EMPTY, __EMPTY_1` |
| 10 | `"", __EMPTY` | `__EMPTY_1, __EMPTY` (header keys first, then lazy) |
| 11 | `__proto__, __proto__` | `__proto__, __proto___1` |
| 12 | `" a", a, "a ", A` | unchanged, all distinct |
| 13 | `" ", " "` | `" ", " _1"` |
| 14 | `x_0, x, x` | `x_0, x, x_1` |
| 15 | `x_01, x, x` | `x_01, x, x_1` |
| 16 | `1, 1` | `1, 1_1` |
| 17 | `a, a_1, …, a_99, a` | …, last is `a_100` (100 probes, once) |
| 18 | 16384 × `a` | `a, a_1, …, a_16383` |
| 19 | 16384 × `""` | `__EMPTY, __EMPTY_1, …, __EMPTY_16383` |

Lazy cases with `make(...)` and the `key` call order (from the tests):

| # | header / data | keys |
|---|---|---|
| 20 | header `a`; row `1,2,3` | `a, __EMPTY, __EMPTY_1` (pinned) |
| 21 | header `a,,c`; row `1`; row `,2` | `named` = `a, c`; column 1 becomes `__EMPTY` on row 2 (pinned) |
| 22 | header missing (`make([])`); cell at column 1 | `__EMPTY` (pinned) |
| 23 | header `a`; row with one cell at column 16383 | `__EMPTY` (pinned, B1) |
| 24 | header `a`; row 1 data at column 2; row 2 at column 1 | column 2 `__EMPTY`, column 1 `__EMPTY_1` |

## 6. Cost bounds and loop progress (AGENTS.md Safety 1, 3, 4)

Implementation shape: `used: Set<string>`, `next: Map<string, number>` (per base,
the next `k` to try, default 1) and `lazy: Map<number, string>`.

`take(b)`: if `b ∉ used`, take `b`. Otherwise start at `k = next.get(b) ?? 1`
and step `k` while `b_k ∈ used`; take `b_k` and set `next(b) = k + 1`.

Why this equals "smallest free `k ≥ 1`": keys are never released. Invariant:
`b_1 … b_{next(b)-1}` are all in `used`.

Why it can't hang, and why the total cost is O(n) expected:
- **Progress:** each iteration strictly increases `k`, and only a key already
  in `used` lets the loop continue. So one call makes at most `used.size + 1`
  iterations. Write this bound in a comment, or assert it with a counter.
- **Amortized:** a string `s` fails a probe for base `b` only if
  `s = b_k`. Since `String(k)` contains no `_`, `b` is `s` minus its last
  `_digits` suffix, so it is unique. And for a fixed `b`, `k` only grows across
  calls. So every key in `used` fails at most one probe in the whole run. Total
  iterations ≤ 2 × (keys handed out). With `n` header entries plus `m` lazy
  columns: O(n + m) expected time (hash operations) and O(total name length)
  memory.
- `m` ≤ the number of distinct columns that carry data cells, which is already
  bounded by `maxCellsPerSheet`. Nothing is allocated per column index, and a
  far cell at column 16383 costs one map entry (Safety 3).
- Key length ≤ raw name length + 1 + 16 (digits of `k` ≤ n + m + 1, far below
  `Number.MAX_SAFE_INTEGER`).
- The hostile inputs are linear: 16384 identical names (every probe succeeds at
  once), 16384 empty headers, and pre-filled suffix ladders (example 17, or
  `a_1 … a_8191` then 8192 × `a`, which is ≤ 16384 failed probes in total).
- No recursion and no regexes, and nothing parses suffixes back out of names.

`make` must not throw on a violated precondition (repeated or descending column
in the entries). If a column repeats, the first entry wins. Every loop stays
bounded regardless.

## 7. Existing tests that must keep passing (unchanged)

From `test/records.test.ts`:
- "duplicate and empty header names are made unique like SheetJS" (example 3)
- "columns past the header get __EMPTY names" (example 20)
- "header row missing from sparse rows: columns get __EMPTY names" (example 22)
- "gap columns inside the header get no defval and are named when data appears" (example 21)
- "defval fills only header-named columns; far cells don't widen every record (B1 regression)" (example 23, `named` only)
- "__proto__ header becomes an own key, not the prototype"
- "dates in header become ISO text; values are cell values" (4.6)
- "explicit header names; extra columns dropped"
- "first-row header", "header by row index skips earlier rows", "defval fills missing header columns; blank rows skipped unless blankRows"
- "B1': a 16384-column header times 1300 rows with defval stays fast, bounded and yielding" (16384 distinct names in `make`)
- "works lazily on streams" (one `RecordKeys` per run)
- the limit tests ("cells count against maxCellsPerSheet, …", "a dense file with defval needs no fills …", "Empty cells count against maxCellsPerSheet", "records built before a limit hit …") must keep their counts: naming keys adds no cells or fills.

New tests to add with the module (`test/recordKeys.test.ts`):
- the table in section 5, row by row;
- hostile timing: 16384 × `a`, 16384 × `""`, and `a_1 … a_8191` + 8192 × `a`,
  each well under 100 ms;
- a fast-check property: for random names drawn from a small alphabet that
  includes `_`, `1`, `""`, `__EMPTY` and `__proto__`, the keys are pairwise
  distinct; each non-empty name `s` maps to `s` or `s_k`; the first occurrence of
  a name that isn't already taken keeps it; the result is deterministic.

## 8. Sources

- PapaParse, MIT. Licence checked:
  <https://github.com/mholt/PapaParse/blob/master/LICENSE> ("The MIT License
  (MIT) Copyright (c) 2015 Matthew Holt"). npm latest is 5.7.0.
  Behaviour only, from:
  - docs <https://www.papaparse.com/docs> (`header` option: duplicated field
    names renamed; `meta.renamedHeaders` example `Column 1_1`);
  - test expectations (inputs/outputs only) in
    <https://github.com/mholt/PapaParse/blob/master/tests/test-cases.js>:
    "Simple duplicated header names" (`A,A,A,A` → `A,A_1,A_2,A_3`),
    "Duplicated header names existing column" (`c,c,c,c_1` →
    `c,c_2,c_3,c_1`), "Duplicate header names with __proto__ field". The first
    case is also in the 5.4.0 tag's tests.
  Used for: the `_k` suffix from 1 (4.1), `__proto__` as an ordinary name
  (4.5), and the rationale (section 2). Not followed: its look-ahead for literal
  names (4.2).
- csv-parse (node-csv), MIT. Licence checked:
  <https://github.com/adaltas/node-csv/blob/master/LICENSE> ("The MIT License
  (MIT) Copyright (c) 2010 Adaltas"). npm csv-parse 7.0.3 declares MIT.
  Docs: <https://csv.js.org/parse/options/columns/> (last value wins on
  duplicate names) and <https://csv.js.org/parse/options/group_columns_by_name/>
  (grouping into arrays). Used for the rationale only (section 2).
- No code from either project was copied or read. Only docs and test
  inputs/outputs.

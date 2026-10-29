# AGENTS.md: effect-spreadsheet rules

Read PLAN.md first. This is the project's rulebook.

## Stack
- Effect v4 only (`effect@^4`). No Effect 3 APIs, no `@effect/*` v3 packages.
- TypeScript strict (incl. `exactOptionalPropertyTypes`), ESM, Node ≥24 for
  dev. The library has to run in Node, in workers, and in browsers: no
  `node:` imports and no native deps under `src/`. Use web-standard APIs
  (`TextDecoder`, `DecompressionStream`, `Uint8Array`/`DataView`).
- `effect` and every `@effect/*` package used at runtime are
  `peerDependencies` (range covering Effect 4, e.g. `^4.1.0`), and also
  `devDependencies` pinned to the version we test against (`4.1.0`). Never in
  `dependencies`: the app's own Effect copy must be the one used, so its
  Data/Schema/Context types match ours. `test/package-deps.test.ts` enforces
  this; every later package follows the same rule.
- Package manager: pnpm. Test runner: vitest + `@effect/vitest`.
  Property/fuzz tests use fast-check.

## Safety (non-negotiable)
1. **Every loop must provably advance.** Each `while`/`for`/recursive
   descent has a cursor that strictly increases, or a counter bounded by a
   limit. Write the progress measure in a short comment when it isn't
   obvious. `continue` must not skip the advance.
2. Hidden rows/columns/sheets are reported (`hidden`/`visibility`). They are
   never a reason to skip work.
3. Cost goes with the cells actually present, never with the declared
   dimension or range. Sparse rows don't allocate dense arrays up to
   the max column.
4. All input is untrusted. Lengths, offsets, counts, sector chains and zip
   sizes are checked against the buffer and against `SpreadsheetLimits`
   before use. Chains are cycle-checked.
5. Bad input → typed `SpreadsheetError` (Schema.TaggedError), never a
   thrown exception/defect and never a hang. Fuzz tests enforce this with a
   time budget.
6. Long work yields: rows are emitted in bounded chunks through `Stream`.

## Code
- Public API in `src/*.ts` modules (`Spreadsheet`, `Workbook`, `Cell`,
  `SpreadsheetError`, `SpreadsheetLimits`) and per-format subpaths. Internals
  go in `src/internal/**`.
- Names have to leave room for writing later: readers are `read*`, and the
  model doesn't depend on any format.
- Keep hot parsing loops in plain imperative TS over `Uint8Array`, and wrap
  them in Effect at the boundaries. Don't allocate an Effect per cell.

## Licensing and SheetJS
- The repo and the npm package are plain MIT (`LICENSE`), with no
  third-party code: no Apache-2.0 code, no `NOTICE`, no `LICENSE-APACHE-2.0`.
- SheetJS is a test oracle and a behaviour reference only, like any other
  project: learn its behaviour (docs, running it, comparing output), but
  never copy its code or closely follow its structure.
- Reference clone: the SheetJS repo in `$SHEETJS_DIR` (default `~/sheetjs`),
  commit d2f2e17.
- SheetJS is never a runtime or package dependency. It may only be loaded
  from `$SHEETJS_DIR` in tests/scripts as an oracle (skip those tests when
  it's absent).
- Corpus fixtures made with the SheetJS writer are generated at script time,
  and the writer's boilerplate parts are stripped from them.

## Privacy
- Never put the user's private projects, companies, customers, hosts or data
  in this repo: not in files, commit messages or PR text. Describe
  motivations generically (e.g. "a hidden-row loop in SheetJS's reader can
  hang a server").

## Tests
- Every format comes with: oracle comparison on corpus files, hidden rows,
  sparse/far cells (e.g. XFD1048576), huge declared dimensions,
  truncated/malformed files, and fuzzing (truncation, bit flips).
- Corpus files are committed only if their license allows it. Everything
  else goes in the gitignored `test/corpus/.cache`. Committed fixtures are
  synthetic (made-up content) and must stay that way.
- `pnpm check` (tsc) and `pnpm test` must pass before any merge.

## Process
- Work on a branch in a git worktree. Main only moves by merges after a
  fresh review.
- No `npm publish` without the user's OK.
- Commit messages: imperative, with the scope (`csv: handle CRLF in quotes`).
- Commit messages, PR text, issues and docs use plain terms: never reference
  internal tracker keys, agent names or review rounds.

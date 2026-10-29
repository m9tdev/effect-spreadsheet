# effect-spreadsheet: plan

An Effect v4 library that reads spreadsheet files with its own pure-TS parsers.
No SheetJS dependency at runtime. Read-only for now, but named and shaped so
writing can come later.

Why: SheetJS 0.20.3 loops forever in `stream.to_json` on hidden
rows of .xls / SpreadsheetML 2003 / SYLK / XLSB (a hidden row does `continue`
before `++R`), which can hang a server that reads such a file.
So the first goal is a reader that can't hang: every loop must provably move forward, and every resource has a limit.

## Formats and order

| Milestone | Scope |
|---|---|
| M0 | Repo, tooling, core model, errors, limits, format sniffing, test harness (SheetJS oracle, corpus, fuzz) |
| M1 | CSV/TSV/delimited text: BOM, UTF-8/UTF-16/Windows-1252, delimiter sniffing, quoting, streaming input |
| M2 | XLSX/XLSM (ZIP + XML): shared strings, inline strings, styles (number formats, so dates can be detected), hidden rows/cols, merges, sheet visibility, defined names (metadata only) |
| M3 | Legacy XLS: CFB container + BIFF8 (BIFF5/BIFF2-4 only if the corpus shows they're needed), codepages, RK/MULRK, SST/CONTINUE, hidden rows |
| M4 | XLSB (ZIP + BIFF12) |
| M5 | SpreadsheetML 2003 XML, then ODS/SYLK/others only if they're needed |

XLSX and XLS are by far the most common formats in practice, so they matter
most. CSV comes first because it's the smallest way to prove the
API, harness and limits end to end.

Not in v1: formula evaluation, charts, images, cell styling beyond number
formats, and encrypted files (they get a typed `Encrypted` error). Full
SSF-style formatted text (`w`) comes later: v1 gives the typed value plus the
number-format code and a date flag. Our own number-format renderer can follow (SSF is a behaviour reference only).

## API sketch (names final, details may move)

```ts
import { Spreadsheet, Workbook, Cell } from "effect-spreadsheet"

// Sniff the format from the bytes (never trust the file name), open the workbook.
Spreadsheet.read(bytes: Uint8Array, options?): Effect<Workbook, SpreadsheetError>
Spreadsheet.readCsv(input: Uint8Array | Stream<Uint8Array>, options?): ...
// Per-format entry points: effect-spreadsheet/csv, /xlsx, /xls, /xlsb

workbook.sheets                     // ReadonlyArray<SheetInfo>: name, index, visibility, dimension
workbook.definedNames               // ReadonlyArray<DefinedName>: name, formula text, scope?, hidden (metadata only)
Workbook.rows(workbook, sheet, opts): Stream<Row, SpreadsheetError>
Workbook.sheetDetails(workbook, sheet, opts): Effect<SheetDetails, SpreadsheetError>
SheetDetails = { hiddenColumns: ColumnSpan[], merges: Range[], dimension?: Range } // one bounded pass, cells skipped
Row  = { index: number, hidden: boolean, cells: ReadonlyArray<Cell> } // cells sparse-safe (column index on each cell)
Cell = Empty | Text | Number | Boolean | Date | Error   (tagged, Schema-defined)
       + optional formula, numFmt
Spreadsheet.toRecords(rows, { header: "first-row" | number | string[] }) // like sheet_to_json
```

- Errors: a `SpreadsheetError` union of `Schema.TaggedError`s:
  `UnsupportedFormat`, `MalformedFile { format, part, offset, reason }`,
  `LimitExceeded { limit, max, actual }`, `Encrypted`. Bad input must never
  throw a defect.
- Sheet layout: hidden columns, merges and the declared dimension come from
  `Workbook.sheetDetails`, not from the rows: in XLSX they sit around the
  cell data (merges after it), so one separate pass over the sheet reads
  them, skipping the cells. Defined names are workbook metadata
  (`workbook.definedNames`); they are not evaluated.
- Limits: a `SpreadsheetLimits` reference with safe defaults, overridable per
  call or by layer. It covers input bytes, decompressed bytes (a hard cap,
  plus an archive-wide inflate budget of max(64 MiB, 200 × archive size)),
  zip entries, sheets, rows, columns, cells per sheet, string length,
  shared-string count, XML depth/attributes, and CFB sector-chain length.
- Streaming: rows come out in bounded chunks, so the event loop yields between
  them (fine in a worker or on the main thread). XLSX/XLS need random access,
  so they take a whole `Uint8Array`. CSV can also take a byte stream.
- Write later: the model (`Workbook`, `Cell`, `Row`) doesn't depend on any
  format. Readers live under `read*`, so `write*`/`Spreadsheet.write` can be
  added without renames.

## Internals (all pure TS, no `node:` imports in `src`)

- `internal/zip`: central-directory reader, inflate via the standard
  `DecompressionStream("deflate-raw")` (Node ≥18, browsers), with a
  byte-counting guard against zip bombs. If that's too slow or too awkward,
  fall back to a small TS inflate.
- `internal/xml`: our own small streaming tokenizer (no DTD/entities beyond
  the predefined five and numeric ones, so no XXE or billion-laughs).
- `internal/cfb`: compound file reader. FAT/miniFAT chains are cycle-checked
  and bounded.
- `internal/biff`: record reader, where every record advances `offset` by
  `4 + len`.
- Text decoding via `TextDecoder` (codepages it supports).

## Safety rules (also in AGENTS.md)

1. Every loop has an explicit progress measure: a cursor that strictly
   increases, or a counter that's bounded by a limit. Tests assert it.
2. Hidden rows/cols are data. They get reported (`hidden: true`), never used
   as a reason to skip advancing.
3. Sparse/far cells (e.g. `XFD1048576`) and huge declared dimensions cost
   memory/time in proportion to the *actual* cells, not the declared range.
4. Malformed or truncated input gives a typed error within bounded time.

## Testing

- vitest + `@effect/vitest`, fast-check for property tests and fuzzing:
  truncation, bit flips, splicing on corpus files. Each case gets a time
  budget, and the outcome must be a value or a typed error.
- Oracle: SheetJS from the local clone `~/sheetjs` (d2f2e17), used only in
  tests and corpus scripts, never as a package dependency. We compare
  `sheet_to_json({header:1, raw:true})`-equivalent output per sheet and
  document intended differences (e.g. hidden-row behaviour, where SheetJS
  hangs).
- Corpus: SheetJS `test_files` (we commit only files whose license allows it,
  and fetch the rest into a gitignored cache), synthetic hidden-row
  fixtures, and generated edge cases (hidden rows, sparse/far cells, huge
  dims, truncation).
- Performance check: 200k-row XLSX, with max event-loop pause measured.

## Packaging

- `effect` (and any `@effect/*` used at runtime) is a peer dependency,
  `^4.1.0`, plus a dev dependency pinned to `4.1.0` for our tests. It is never
  in `dependencies`, so apps don't end up with two Effect copies and
  mismatched Data/Schema/Context types. A test fails if `effect` or
  `@effect/*` appears in `dependencies`.

## Licensing

The repo and the npm package are plain MIT, with no third-party code: no
Apache-2.0 code, no `NOTICE`, no `LICENSE-APACHE-2.0`. SheetJS is a test
oracle and a behaviour reference only: we learn its behaviour, but never
copy its code or closely follow its structure. It is loaded only in tests
and corpus scripts, from `$SHEETJS_DIR` (default `~/sheetjs`). Corpus
fixtures made with the SheetJS writer are generated at script time, with
the writer's boilerplate parts stripped.

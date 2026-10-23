# Test corpus and SheetJS oracle

Everything here is test/script code. Nothing under `scripts/corpus` or
`test/corpus` is imported from `src/`, and SheetJS is never a package
dependency: the scripts load it from the local clone `~/sheetjs` (d2f2e17,
SheetJS 0.20.3). The convention is shared with the test-side oracle: env
`SHEETJS_DIR` is the SheetJS **repo directory**, default `path.join(os.homedir(), "sheetjs")`
(an empty value counts as unset),
and each side picks its own entry file inside it. These scripts use `xlsx.js`
(`SHEETJS_ENTRY` in `lib.mjs`).

The scripts are plain Node ESM (Node ≥ 24), with no dependencies.

```sh
node scripts/corpus/fetch.mjs          # SheetJS test_files -> test/corpus/.cache/sheetjs (pinned, idempotent; --timeout SEC, default 900)
node scripts/corpus/generate.mjs       # our fixtures -> test/corpus/generated (deterministic)
node scripts/corpus/generate.mjs --check
node scripts/corpus/hidden-rows/generate.mjs --check   # hidden-row fixtures regenerate identically (zips by content)
node scripts/corpus/hidden-rows/manifest.mjs    # test/corpus/hidden-rows/manifest.json
node scripts/corpus/manifest.mjs       # full manifest -> test/corpus/.cache/manifest.json (gitignored)
node scripts/corpus/oracle.mjs         # expected output -> test/corpus/.cache/expected/<sha256>.json
node --test scripts/corpus/*.test.mjs   # unit tests (BOF sniffing, quirk checks, expected.mjs cleanup), no SheetJS needed
node scripts/corpus/oracle-selftest.mjs   # oracle process hygiene (orphans, startup deaths, tooBig docs), ~15 s
node scripts/corpus/expected.mjs       # committed SheetJS expectations for generated + hidden-rows (after oracle --source generated/hidden-rows)
node scripts/corpus/expected.mjs --check
```

Oracle options: `--force`, `--jobs N` (default min(10, cores/2)), `--timeout SEC` (parse,
default 60), `--probe-timeout SEC` (default 15), `--no-probe`, `--source sheetjs|generated|hidden-rows`,
`--format xls,xlsb`, `--max-size BYTES`, `--limit N`, `--ready-timeout SEC` (default 30), `--verbose`. It skips files whose expected
JSON already exists (same schema) unless you pass `--force`. A run report goes to
`test/corpus/.cache/oracle-report.json`. If you run with a filter (`--source`), the report only
covers that run. Run without filters at the end: it reuses the cache and writes the full report.

## Layout

| path | committed | what |
| --- | --- | --- |
| `scripts/corpus/lib.mjs` | yes | zip/CFB/BIFF readers that are only good enough to classify files; byte sniffing; feature detection |
| `scripts/corpus/fetch.mjs` | yes | download + extract the SheetJS corpus |
| `scripts/corpus/manifest.mjs` | yes | builds the full manifest `test/corpus/.cache/manifest.json` |
| `scripts/corpus/oracle.mjs` | yes | SheetJS oracle (pooled child processes, timeouts) |
| `scripts/corpus/oracle-selftest.mjs` | yes | oracle process-hygiene tests (orphaned children, startup deaths) |
| `scripts/corpus/expected.mjs` | yes | exports committed expectations for our fixtures to `test/corpus/expected/` |
| `scripts/corpus/generate.mjs` | yes | our edge-case fixtures |
| `scripts/corpus/hidden-rows/*.mjs` | yes | hidden-row fixture generator (also writes `renamed/`); hidden-rows manifest; shared names |
| `test/corpus/sources.json` | yes | pin of the external corpus: URL, archive sha256/size, file count (no per-file details) |
| `test/corpus/generated/` | yes | our fixtures (~170 KB; size target ~200 KB), with `README.json` (name, sniffed format, size, sha256, description) |
| `test/corpus/hidden-rows/` | yes | Synthetic hidden-row fixtures + `manifest.json` |
| `test/corpus/expected/{generated,hidden-rows}/<file>.json` | yes | SheetJS expected output for our fixtures, so CI can compare without `~/sheetjs` (58 files, ~152 KiB, each ≤ 64 KiB) |
| `test/corpus/sheetjs-quirks.json` | yes | list of places where that SheetJS output is wrong or differs on purpose; every `sheetjs` and `correct` value is checked (`scripts/corpus/quirks.mjs`) |
| `test/corpus/.cache/` | **no** (gitignored) | `sheetjs/test_files` (fetched, 1.9 GB), `manifest.json` (full manifest), `expected/` (oracle output, ~1.3 GB), `oracle-report.json`; about **3.2 GB** in all (apparent size; plus 341 MB with `fetch --keep-zip`) |

**Nothing committed lists third-party corpus files.** GitHub blocked `SheetJS/test_files`
for `private_information`, and we may not redistribute much of it, so file names, paths
and per-file details of that corpus stay in the gitignored cache. The committed metadata
covers only our own files (`generated/README.json`, `hidden-rows/manifest.json`) plus the archive
pin (`sources.json`). The docs name a few SheetJS-origin files (e.g. `AutoFilter.*`,
`comments_stress_test.*`) as examples.

## SheetJS test_files: source and licence

**Source.** SheetJS keeps its test corpus out of the main repo. `~/sheetjs/Makefile`
(`make init`) downloads `https://test-files.sheetjs.com/test_files.zip` and
`tests.lst` lists the files its tests use. GitHub blocks the GitHub repo
`SheetJS/test_files` ("Repository access blocked", reason `private_information`,
2024-07-19), and git.sheetjs.com has no public `test_files` repo (the
README's `…/sheetjs/test_files/releases` link gives 404). So there is no commit
to pin. `fetch.mjs` pins the archive (in `test/corpus/sources.json`) by **sha256
`a786d4f1a80adf9447997b6f3b30b2593c776c27e0b075e3971627c17205f220`** (341,452,051
bytes, Last-Modified 2024-07-17). That is 12,237 files, about 1.9 GB unpacked. If
upstream changes the archive, the script stops; `--update` accepts the new one and rewrites `sources.json`.
The download and extraction take about 10 s here.

**Licence** (from `test_files/README.md` and `LICENSE`, Apache-2.0):

- The README says the root files are Apache-2.0, as are the Excel re-saves in the
  201*/ and biff*/ directories of originals. We treat a file as **SheetJS-original
  (Apache-2.0)** only if two things hold. First, it is at the root or directly in a 201*/
  or biff*/ directory. Second, its name has either no `<prefix>_` or a stem that the README
  documents (`SHEETJS_STEMS` in `manifest.mjs`, e.g. `AutoFilter`, `comments_stress_test`,
  `number_format`).
- The **other SheetJS subdirectories** (`1904/`, `crlf/`, `dpi/`, `wps/`, `artifacts/`,
  `numbers/`, `duration_112/`, `datetime_8192/`, `issue*/`, …), and names with an **unknown
  prefix**, are `unclear` and not committable. The README makes no licence statement for them.
- **External** files carry a source prefix in their name or sit in a clone directory
  of the source project. A derivative keeps the licence of its original. The README
  gives these licences ("at the time of this writing"). `manifest.mjs` holds the prefix
  rules (`ORIGINS`):

| origin | licence | files in manifest | committable? |
| --- | --- | --- | --- |
| SheetJS originals (root + 201*/biff* re-saves, known stems) | Apache-2.0 | 866 | **no** (repo is MIT only) |
| other SheetJS subdirectories, unknown prefixes | unclear | 126 | **no** |
| Apache POI | Apache-2.0 | 4299 | **no** (repo is MIT only) |
| spout | Apache-2.0 | 118 | **no** (repo is MIT only) |
| OpenOffice (oo34xml) | Apache-2.0 | 36 | **no** (repo is MIT only) |
| openpyxl | MIT | 121 | yes (copyright notice) |
| roo | MIT | 747 | yes (copyright notice) |
| Spreadsheet::ParseXLSX | MIT | 172 | yes (copyright notice) |
| xlrd | BSD-3-Clause | 84 | yes (copyright notice) |
| Excel::Reader::XLSX | Perl Artistic | 47 | **no** |
| pyExcelerator | BSD-4-Clause (advertising clause) | 33 | **no** |
| jxls | LGPL | 697 | **no** |
| LibreOffice | LGPL-3.0 per README | 642 | **no** |
| PHPExcel | LGPL | 2 | **no** |
| OpenOffice test docs (ootest) | unclear (README: "license terms are unclear") | 116 | **no** |

Each entry of the (cached) full manifest has `origin`, `license` and `committable`. The
repository is plain MIT, so nothing of Apache-2.0 origin is committable. Only the MIT and
BSD-3 origins could ever be committed, with their copyright notice.

**Decision: we commit no SheetJS corpus files.** They are fetched (pinned) into
`test/corpus/.cache`. Fetching is fast and reproducible. Committing ~1.9 GB makes
no sense. The README also says the third-party test sets may hold files their
contributors had no right to release ("Requests for Removal"). The archive itself
is no longer browsable upstream, so we don't publish its file listing either. If we
ever need a specific file in git (e.g. for CI without network), pick it from a
`committable: true` origin (MIT or BSD-3), and keep its copyright notice with it. Better
still, generate an equivalent fixture of our own.

The CSV/JSON/XML files named `<book>.<ext>.<n>.csv` / `<book>.<ext>.xml` are Excel or
SheetJS exports of the workbook next to them. They have `derived: true` in the manifest.
Most of the corpus CSVs are Excel-exported, so they are real CSV input. SheetJS's own curated list (`~/sheetjs/tests.lst`) gets `sheetjsTests: true`.

## Full manifest (`test/corpus/.cache/manifest.json`, gitignored)

`manifest.mjs` rebuilds it from what is on disk (about 6 s). `oracle.mjs` reads it, so
run `fetch` and `manifest` first. Without the fetched corpus, it still lists the generated
and hidden-rows files.

```jsonc
{
  "version": 1,
  "sources": { "sheetjs": {"root": "test/corpus/.cache/sheetjs/test_files", "archiveSha256": "…"},
               "generated": {"root": "test/corpus/generated"}, "hidden-rows": {"root": "test/corpus/hidden-rows"} },
  "summary": { "xls/biff8": {"files": 1213, "unique": 676, "bytes": …, "hiddenRows": 59, "encrypted": 21, "generated": 4, "hidden-rows": 2}, … },
  "excludedFormats": { "lotus": 33, … },          // SheetJS-corpus files of non-focus formats (not listed)
  "files": [                                        // one JSON object per line
    {"source":"sheetjs","path":"AutoFilter.xlsx","format":"xlsx","size":…,"sha256":"…",
     "origin":"sheetjs","license":"Apache-2.0","committable":false,"sheetjsTests":true,
     "features":{"entries":…,"mainPart":"xl/workbook.xml","sheets":…,"hiddenRows":…}}
  ]
}
```

The **format** comes from the bytes, never the extension (`lib.mjs` `sniff`):

- ZIP: the `officeDocument` target in `_rels/.rels` and its content type decide
  between xlsx/xlsm/xltx/xlsb. A bare match on `[Content_Types].xml` isn't enough,
  because SheetJS writes a `Default Extension="bin"` xlsb type into every xlsx.
  A `mimetype` entry means ods; `Index/*.iwa` means Numbers.
- CFB: `EncryptionInfo`/`EncryptedPackage` → `ooxml-encrypted`. A `Workbook`/`Book`
  stream → `xls`, with `biff` taken from its BOF.
- A raw BIFF2–5 BOF → `xls` with `container: "raw"`.
- Text: SYLK (`ID;`), DIF, SpreadsheetML 2003 (`urn:schemas-microsoft-com:office:spreadsheet`
  plus `<Workbook>`) → `xlml`, flat ODS → `fods`, HTML, JSON, other XML. Anything else that
  looks like text → `csv`, with `encoding`, `bom`, `delimiter`, `lineEnding` and
  `quotedNewline` features.

The **focus formats** listed are csv, xlsx, xlsm, xltx, xls, xlsb, xlml,
ods, fods, sylk and ooxml-encrypted (`--all-formats` lists the rest too). Generated and hidden-rows
files are always listed. **Features** are best effort and cheap: `biff`, `encrypted`
(FILEPASS), `hiddenRows` (a count), `sheets`/`hiddenSheets`/`veryHiddenSheets`,
`hugeDimension` (a declared range that reaches the format's last row or column: xlsx
`<dimension>`, xlsb BrtWsDim, xls DIMENSIONS, xlml ExpandedRow/ColumnCount, SYLK `B;`), `vba`, `truncated` (BIFF record past the end of the stream) and `partial`
(a part was too big to scan).

### Counts per format

Counts from the manifest at the time of writing (sheetjs + generated + hidden-rows; *unique* = distinct sha256):

| format | files | unique | with hidden rows | encrypted | generated | hidden-rows |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| csv | 4324 | 1962 | – | – | 12 | 0 |
| xlsx | 1019 | 564 | 33 | – | 10 | 4 |
| xlsm | 45 | 27 | 4 | – | 0 | 0 |
| xltx | 2 | 1 | 0 | – | 0 | 0 |
| xls BIFF8 | 1213 | 676 | 59 | 21 | 4 | 2 |
| xls BIFF5 | 55 | 49 | 2 | 3 | 2 | 0 |
| xls BIFF4 | 7 | 6 | 0 | 0 | 1 | 0 |
| xls BIFF3 | 5 | 4 | 0 | 0 | 1 | 0 |
| xls BIFF2 | 3 | 3 | 0 | 0 | 1 | 0 |
| xlsb | 504 | 494 | 18 | – | 4 | 2 |
| xlml (SpreadsheetML 2003) | 831 | 773 | 36 | – | 4 | 2 |
| ods | 116 | 82 | 3 | – | 2 | 0 |
| fods | 4 | 4 | 0 | – | 1 | 0 |
| sylk | 25 | 24 | 7 | – | 8 | 2 |
| ooxml-encrypted | 15 | 9 | – | 15 | 0 | 0 |
| empty (0 bytes) | 1 | 1 | – | – | 1 | 0 |

(The sheetjs `empty` files are 2711 empty Excel CSV exports of empty sheets. They are
counted under `excludedFormats`, and generated `csv_empty.csv` stands in for them.)

## Oracle output (`test/corpus/.cache/expected/<sha256>.json`)

There is one file per unique sha256. `oracle.mjs` runs SheetJS in pooled child
processes (`--max-old-space-size=4096`), with a timeout per file. A timeout kills
the child (SIGKILL) and gets recorded. SheetJS **mutates the options object you give
`XLSX.read`**, so a long-lived worker must pass a fresh copy each time. Before this
was fixed, most files failed with `x.charCodeAt is not a function`.

Process hygiene: the parent tracks its live children and SIGKILLs them on exit,
SIGINT, SIGTERM and SIGHUP. Each child runs a watchdog `worker_thread` that SIGKILLs
its own process when its parent pid changes (the parent was killed alone) or when the
current job's hard deadline passes (job timeout + 10 s). The watchdog thread keeps
running while the main thread spins in a SheetJS loop. A child that dies or doesn't
say `ready` within `--ready-timeout` counts as a startup death. After 3 consecutive
startup deaths the oracle aborts with exit 2, instead of respawning forever (e.g. when
`xlsx.js` throws). `oracle-selftest.mjs` checks these cases in a temporary cache
(`CORPUS_CACHE_DIR`):
- the parent SIGKILLed alone;
- the parent SIGTERMed (its handler must exit 143, since the watchdogs alone would also
  clean up);
- a throwing `xlsx.js`;
- an `xlsx.js` that never finishes loading;
- `tooBig` results under a tiny `ORACLE_MAX_CELLS`, which must keep `reason`, `options`,
  `workbook`, per-sheet counts and the SheetJS version.

Cases that need SheetJS are skipped, not failed, when it is missing. Temp dirs left by
a SIGKILLed selftest are removed at the next start. The parent only writes the expected
file itself for `timeout` and `crash`; `ok`, `error` and `tooBig` documents come from the
child.

SheetJS calls:

```js
XLSX.read(buf, { type: "buffer", dense: true, cellDates: false, cellNF: true, cellFormula: true,
                 cellText: true, cellStyles: true, sheetStubs: false, WTF: false })
XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true, UTC: true })  // only when !ref area <= 2,000,000
```

Why: `dense` keeps memory proportional to the cells that are present. `cellDates: false`
keeps dates as serial numbers, plus `z`. `cellStyles: true` is **required** for SheetJS's
xlsx reader to put hidden rows in `!rows`; without it the flag is silently dropped. In
`rows`, `sheet_to_json` still turns date-formatted numbers into Dates. `UTC: true`, and
children running with `TZ=UTC`, keep those (ISO strings in the JSON) independent of the
machine's time zone. Schema 1 files had local-time Dates. We never
use `XLSX.stream.to_json` for expected output, because that is where the hang is (below).

```jsonc
{
  "schema": 2, "oracle": "sheetjs", "sheetjsVersion": "0.20.3", "sheetjsCommit": "d2f2e179636b",
  "sha256": "…", "path": "AutoFilter.xlsx", "source": "sheetjs", "format": "xlsx", "kind": "parse",
  "ms": 12,
  "status": "ok",                      // "ok" | "tooBig" | "error" | "timeout" | "crash"
  "error": "…",                        // only if not ok
  "options": { "read": {…}, "json": {…}, "rowsMaxArea": 2000000 },
  "workbook": {
    "bookType": "xlsx", "date1904": false,
    "sheets": [ { "name": "Sheet1", "visibility": "visible" } ]   // "visible" | "hidden" | "veryHidden"
  },
  "sheets": [ {
    "name": "Sheet1",
    "ref": "A1:C4",                    // SheetJS !ref (null for an empty sheet); fullref if it differs
    "hiddenRows": [2],                 // 0-based row indices with !rows[i].hidden
    "hiddenCols": [],                  // 0-based column indices with !cols[i].hidden
    "merges": ["A1:C1"],
    "cellCount": 12,
    "cells": [                          // sparse, row-major; every cell SheetJS produced
      { "r": 0, "c": 0, "t": "s", "v": "Product" },
      { "r": 1, "c": 1, "t": "n", "v": 46024.52, "z": "yyyy-mm-dd", "w": "2026-01-02" },
      { "r": 1, "c": 2, "t": "n", "v": 2, "f": "B2*2" }
    ],                                  // t: s n b e d (z only when not General; w = SheetJS-formatted text; F = array range)
    "rows": [ ["Product","Qty","Price"], … ]   // sheet_to_json header:1 — starts at !ref's first row/col!
                                       // null + "rowsOmitted" when the !ref area > 2,000,000 cells
  } ]
}
```

**Caps.** A workbook with more than 5,000,000 cells (`MAX_CELLS`), or an expected file that
would exceed 128 MiB (`MAX_BYTES`), gets `status: "tooBig"` with a `reason`, the
`workbook` block, and per-sheet counts only (`name`, `ref`, `cellCount`, `hiddenRowCount`).
Before the caps, the cache held about 32 GB of expected JSON. One shared-strings bomb from
a third-party set alone produced 25 GB. With the caps it's about 1.3 GB, and the largest
file is about 116 MB.

Non-finite numbers are written as `{"$num":"NaN"}`. Big results are written in a
streaming way, because some sheets exceed V8's maximum string length.

### stream.to_json probe (the hidden-row hang)

The oracle also runs SheetJS's streaming reader the way a typical SheetJS-based
reader does, on every file that has hidden rows (manifest `features.hiddenRows`) and on every
generated/hidden-rows file. That gives three jobs per file, each in its own child, with a 15 s
kill. The results go to `expected/<sha256>.stream.<variant>.json`:

| variant | read | stream.to_json options |
| --- | --- | --- |
| `default` | `XLSX.read(buf)` (typical reader) | `{raw:true, skipHidden:true, UTC:true, header:"A"}` |
| `cellStyles` | `+ cellStyles:true` | same |
| `control` | `+ cellStyles:true` | same but `skipHidden:false` |

`oracle-report.json` → `streamVerdicts` gives each file a verdict:

- **hang**: `default` or `cellStyles` timed out and `control` finished. This is the
  `skipHidden` loop.
- **slow**: `control` timed out too. That is a range scan over a huge `!ref`, not the
  hidden-row loop.
- **ok** / **error**.

Bug mechanism (`~/sheetjs/xlsx.js` ~27899, `write_json_stream` `_read`): the loop is
`while (R <= r.e.r) { if ((rowinfo[R-1]||{}).hidden) continue; … ++R; }`. The `continue`
skips `++R`, so the loop spins forever. With `header:1`/`"A"` the offset is 0, so the
loop is stuck on the row **after** each hidden row. A hidden row that is the last row
of the range doesn't hang (see `hidden_rows_last.slk`). It only matters when `XLSX.read`
fills `!rows`. That happens always for xls/xlsb/xlml/SYLK, and for xlsx only with
`cellStyles:true`. ODS/FODS never record hidden rows on read.

### Results (2026-10-23, `oracle.mjs --jobs 8 --force`, 460 s wall)

**Parse** (one job per unique sha256, 60 s limit each): 4674 jobs. 4628 are ok and 19
are `tooBig` (counts only: the large-strings, calendar-stress and datetime_8192 sets and
the million-row files). None hit the parse timeout and none crashed. That "0 timeouts" is
only for the parse step: every timeout of the run is in the stream probe below (211 probe
runs, by design). There are 27 errors. All of them are inputs that are meant to fail:
- 18 password-protected xls (SheetJS's own `password_*` set and third-party ones)
- 1 truncated BIFF record
- 6 corrupt or unusual zips (missing parts, bad sizes, an encrypted ods)
- 1 ods with an invalid duration
- 1 XML file that isn't SpreadsheetML
- 1 SheetJS string-length limit on a very large ods

The 9 unique `ooxml-encrypted` files are skipped.

**stream.to_json probe**: 156 files → **108 hang**, 6–7 slow, 41–42 ok. Every hang has
`control` (skipHidden:false) finishing, which confirms the hidden-row loop:

| format | hang | read that hangs | examples (SheetJS-origin / ours) |
| --- | ---: | --- | --- |
| xls (BIFF5/8) | 35 | default | `AutoFilter.xls`, `comments_stress_test.xls` (BIFF8 and BIFF5), `generated/hidden_rows{,.biff5}.xls`, `hidden-rows/hidden-row.xls` |
| xlml | 29 | default | `AutoFilter.xml`, Excel re-saves of `comments_stress_test`, `generated/hidden_rows.xml`, `hidden-rows/hidden-row.xml` |
| xlsb | 17 | default | `AutoFilter.xlsb`, `comments_stress_test.xlsb`, `wtf_path.xlsb`, `generated/hidden_rows.xlsb`, `hidden-rows/hidden-row.xlsb` |
| sylk | 6 | default | `generated/hidden_rows_{first,middle,consecutive,all}.slk`, `generated/sylk_row_format_no_m.slk` (SheetJS's false hidden row), `hidden-rows/hidden-row.slk` (`hidden_rows_last.slk` does **not** hang, as the `R-1` analysis predicts) |
| xlsx/xlsm | 21 (19 + 2) | **only `cellStyles:true`** | `AutoFilter.xlsx`, `comments_stress_test.xlsx`, `wtf_path.xlsx`, `generated/hidden_rows.xlsx`, `generated/many_empty_rows.xlsx`, `hidden-rows/hidden-row.xlsx` |

The rest of the 108 are third-party corpus files; they are listed only in the cached report.
The slow ones (control times out too: a huge `!ref`, not the loop) are
`generated/sparse_far.{xlsx,xlsb,xml,slk}` and `generated/huge_dimension.{xlsx,xlsb}`.
`generated/sparse_far.xls` (a 16M-cell range) is borderline: it is "ok" on an idle machine
and "slow" under load, because its scan takes about the 15 s probe timeout.

The full list is in `test/corpus/.cache/oracle-report.json` → `streamVerdicts`.

### Committed expectations for our fixtures (`test/corpus/expected/`)

`expected.mjs` copies the cached oracle result for each `generated/*` file and each of
the 6 hidden-rows originals to `test/corpus/expected/<source>/<file>.json`. It drops `ms`. It leaves out the stream-probe
statuses, because whether a probe counts as "slow" depends on machine load. The hang facts
for hidden-rows are in `hidden-rows/manifest.json` (`oldReaderHangs`). hidden-rows `renamed/` copies
map to their original via `hidden-rows/manifest.json` `copyOf`. Each file is capped at 64 KiB:
over the cap, `rows` is dropped first (`rowsOmitted: "size cap"`), then `cells`, while
counts, hidden rows, merges and refs stay. At the time of writing, only
`many_empty_rows_implicit.xlsx` hits the cap (1,048,576 dense rows). `--check` regenerates
the set into a temp dir and compares it both ways.

## Hidden rows: our reader's contract

**Our reader keeps hidden rows and reports them with `hidden: true`. It never
skips them** (AGENTS.md safety rule 2). For `test/corpus/hidden-rows/hidden-row.xlsx`, the
expected result is **4 rows, one of them hidden** (index 2, "Product B"). That is
the same for every `hidden-row.*` file. `plain.xlsx` has 4 rows, none hidden.
Callers who want SheetJS's `skipHidden` behaviour filter on `row.hidden`
themselves. This is a deliberate difference from SheetJS, which drops hidden rows
(and hangs in the stream reader). The expected values for these files are in
`test/corpus/hidden-rows/manifest.json` (`expected.sheets[].rows[] = {index, hidden, values}`).

## Generated fixtures (`test/corpus/generated/`)

`generate.mjs` writes them. The output is deterministic (`--check` regenerates it into
a temp dir and compares). A caveat for zips (xlsx, xlsb, ods): deflate output depends on the
zlib build. Node's bundled zlib and a system zlib, or different CPUs, can compress the same
data differently. So `--check` lets zips differ only in what compression changes: each entry's
method, compressed size, compressed data, and therefore the offsets (`zipDiff` in `lib.mjs`).
Everything else must be equal: every other local and central header field (timestamps, CRC32,
sizes, flags, attributes, names, extra fields, comments), the entry order, the archive comment
and any trailing bytes. Each entry's CRC32 must also match its uncompressed bytes. `--check`
reports how many zips differ only in compressed bytes; any other difference is a DIFF, with
the reason. Every other file, and the committed `README.json` against the committed bytes,
must match exactly. Zips are rebuilt with fixed timestamps, and document properties
are fixed (author `effect-spreadsheet`, 2026-01-01). They are our own files (MIT). They
are generated with SheetJS 0.20.3 at script time (never a dependency). Where the SheetJS
writer can't express a case, they are built or patched by hand.

The writer's own boilerplate is stripped from every fixture (`stripWriterBoilerplate` in
`lib.mjs`). The cell and oracle output is unchanged: cells with t/v/w/z/f, refs, hidden
rows, merges and `rows` in `test/corpus/expected/` are identical except for the files'
`sha256`. What does
change is what SheetJS derives from the stripped parts: the theme-resolved font colours
(`rgb` in its style objects) and the application name. The oracle records neither. The
stripped parts are:
- xlsx/xlsb: `xl/theme/theme1.xml` with its relationship and Content_Types override,
  `<Application>SheetJS</Application>` in `docProps/app.xml`, and the app/version strings in
  the xlsb BrtFileVersion record (now empty).
- ods/fods: the `meta:generator` element and the `sh33tjs` namespace.
- SYLK: `ID;PSheetJS;N;E` becomes `ID;PWXL;N;E`.
- xls: the `\x01Sh33tJ5` CFB stream is unlinked and zeroed. The WRITEACCESS user name is
  blanked, and so is the property-set SystemIdentifier.

SheetJS's BIFF writers drop row properties, and its writers iterate the whole declared
range, so far cells or huge dimensions would take forever. Those cases are built or patched
by hand. Each file is described in `generated/README.json`.

- **Hidden rows** (sheets `first`, `middle`, `last`, `consecutive`, `all`; 6 rows × 2 cols)
  in `hidden_rows.{xlsx,xlsb,xml,xls,biff5.xls,ods}`. The xls/biff5 files have ROW records
  (fDyZero) inserted by hand, with BOUNDSHEET offsets fixed up. The ods file has
  `table:visibility="collapse"` added by hand. There are also SYLK files
  `hidden_rows_{first,middle,last,consecutive,all}.slk` (`F;M0;R<n>`), and
  `sylk_row_format_no_m.slk` (a row-format record without `M`, which is not a hidden row).
- **Sparse/far cells** `sparse_far.{xlsx,xlsb,xml,slk}`: A1 and XFD1048576.
  `sparse_far.xls`: A1 and IV65536 (the BIFF8 maximum).
- **Huge declared dimension**, 2 cells: `huge_dimension.xlsx` (`<dimension ref="A1:XFD1048576"/>`),
  `.xlsb` (BrtWsDim), `.xls` (DIMENSIONS A1:IV65536), `.xml` (ExpandedRow/ColumnCount), `.slk` (`B;Y1048576;X16384`).
- **Many empty rows**: `many_empty_rows.xlsx` has 5000 `<row r/>`, some hidden or with a custom height.
  `many_empty_rows_implicit.xlsx` has 1,048,574 `<row/>` without `r`, then a cell in row 1048576
  (about 6 MB of XML in an 11 KB zip). `many_empty_rows.csv` has 5000 blank lines.
- **Strings** `strings.xlsx`: shared strings with rich runs, entities, `_xHHHH_`, `xml:space`,
  unicode, an empty string and phonetic runs; inline strings; `t=str/b/e/d`; rows and cells without `r`.
- **Formulas** `formulas.xlsx`: plain, shared (`si`) and array formulas; string, error and bool
  cached results; one formula with no cached value.
- **Dates** `dates.xlsx` / `dates_1904.xlsx`: built-in formats 14/15/17/20/22/45–47 and custom
  formats (locale prefix, quoted text, `[h]`, a non-date `0.00`) on serials 46024.52, 0, 60
  (the 1900 leap bug), 61, 1.75 and -1.
- **Features workbook** `features.{xlsx (SST), xlsb, xls, biff5.xls, xml, ods, fods, slk}`: sheets Data
  (text, numbers, booleans, unicode, formulas with cached values, errors), Dates, Merges, Hidden,
  VeryHidden, and a unicode sheet name. BIFF5 can't store the unicode name. Also
  `basic.biff{2,3,4}.xls` (raw BIFF streams).
- **CSV**: UTF-8 BOM, `;` with decimal comma, UTF-16LE tab ("Unicode Text"), UTF-16BE, quoted
  LF/CRLF and `""`, CR-only, Windows-1252, ragged rows, no trailing newline, empty, BOM-only, TSV.

**Truncated / bit-flipped variants are not committed.** Tests derive them at test time from
these files (truncate at every N bytes, flip bits) with fast-check, under a time budget.

## Hidden-row fixtures (`test/corpus/hidden-rows/`)

These are synthetic files that reproduce the SheetJS hidden-row hang. All content is
made up: `Sheet1`, 4 rows (Product/Qty/Price; Product A/B/C), and row index 2 is
hidden in `hidden-row.*`. The generator
`scripts/corpus/hidden-rows/generate.mjs` uses `~/sheetjs` 0.20.3, strips the writer's
boilerplate (see the generated fixtures below), and reproduces all six byte for byte
(`--check`). `renamed/` holds byte copies with misleading extensions (xls as `.csv`,
xml as `.xlsx`, slk as `.csv`, xlsb as `.xls`, xlsx as `.csv`, plain.xlsx as `.xls`).
These carry `copyOf` and `extensionMismatch: true` and prove that detection ignores the
name. Tests must use the repo paths only.

| file | sha256 | hidden row stored as | old reader hangs (default read / cellStyles) |
| --- | --- | --- | --- |
| hidden-row.xls | `2fd422228fc4443613ec146cd15e3ccb30c904f28e0a5ffd067389188d8514cd` | BIFF8 ROW 0x0208 fDyZero | yes / yes |
| hidden-row.xml | `479174c732d3ac9a2b19c719b0ffa64d86f4a4f74dcadf73e24a107d2b6ab9d0` | `ss:Hidden="1"` | yes / yes |
| hidden-row.slk | `e6dd697845034d1b4805f4e7062ad9363ebef5b7098f937c6a3e10ab657fd971` | `F;M0;R3` | yes / yes |
| hidden-row.xlsb | `8d94da67eddb477e1066fc0b9dd7a647ed87f6dd9a515a31301b1348fd3b7f68` | BrtRowHdr 0x10 | yes / yes |
| hidden-row.xlsx | `e45f30a52f3796c19a4e136e929ddb99c92a5da4008f10fc2c581891f9e62c67` | `<row hidden="1">` | no (flag lost) / yes |
| plain.xlsx | `b785b1b4d4a796296242cebfdab3fa09dc889393d28cb60d2d812da26d73d458` | – | no / no |

Our oracle probe confirms every entry in that table (verdict `hang`). The metadata check
was done before committing; the findings are in the commit message of the hidden-rows import.
Summary: every document property is empty. Zip entries carry no timestamps (1980-00-00),
comments or extra fields. The CFB has no timestamps. WRITEACCESS is blank (only spaces). `strings`
(ASCII and UTF-16LE, also on the unzipped parts and the Workbook stream) shows no names,
user names, hostnames, e-mails, company names or paths.

## Known SheetJS quirks seen through the oracle (intended differences)

The quirks that affect our committed fixtures are listed in `test/corpus/sheetjs-quirks.json`.
Tests compare against the SheetJS expectation except at those places. Each entry has an `id`,
`file`, `sheet`, `field`, the `sheetjs` value, the `correct` value, and a `correctCheck` method.
Value formats depend on `field`:
- `hiddenRows`: an array of row indices.
- `ref`: an A1 range.
- `cells`: an array of `[r, c, v]` triples. `v: null` means "no cell", e.g.
  `csv-empty-quoted-field-no-cell` has `sheetjs: [[5, 1, null]]`.
- `dimension`: `sheetjs` is `{ref}` (SheetJS's `!ref`); `correct` is `{cells, declaredRows, declaredCols}` (the span of the cells and the declared size).

Every value is checked; none is marked unchecked. `expected.mjs` (through `quirks.mjs`)
checks each `sheetjs` value against the oracle result. It checks each `correct` value against
the fixture bytes:
- `sylk-hidden-rows`: the `F;…M0…;R<n>` records.
- `csv:<encoding>`: the bytes decoded with `TextDecoder` and split by a small RFC 4180 reader.
- `xlsx-shared-strings`: the shared strings, with entities and `_xHHHH_` decoded.
- `declared-dimension`: the sniffed declared range, plus the span of the cells.

A mismatch fails the export. `expected.mjs` also stamps the ids into the affected expected
files (`knownQuirks`). `node --test scripts/corpus/quirks.test.mjs` checks that every quirk
passes, and that a wrong `sheetjs` or `correct` value fails for each field type and check
method.

- **xlsx hidden rows need `cellStyles:true`.** Otherwise `!rows` is empty.
- **SYLK false hidden rows.** In `F;…;R<n>` without `M`, SheetJS keeps `Mval = 0` and
  marks the row hidden (`bits/40_harb.js` ~563). See `generated/sylk_row_format_no_m.slk`:
  `F;SDM6;R1` is a row-style record (row 1 not hidden), `F;M0;R3` hides row 3. A row is
  hidden only with an explicit `M0`.
- **UTF-16BE CSV** isn't decoded (garbage cells). See `generated/csv_utf16be.csv`.
- **Windows-1252 CSV** is decoded as ISO-8859-1: bytes 0x80–0x9F become C1 controls
  instead of €, curly quotes and so on (`generated/csv_windows1252.csv`).
- **Astral character references** (`&#x1F363;`) are truncated to 16 bits (U+F363), and
  **`_x000D_`** is dropped (`generated/strings.xlsx`).
- An explicit empty quoted CSV field (`""`) gives no cell (`generated/csv_quoted_newlines.csv`).
- **BIFF5 sheet names** lose non-Latin-1 characters on write. That is a property of
  `features.biff5.xls`, not a reader bug.
- SheetJS ignores declared dimensions in xlml/SYLK (`ref` = the actual cells), but
  uses them as-is for xlsx/xlsb/xls (`A1:XFD1048576`).
- `stream.to_json` cost goes with `!ref`, so huge-dimension files are "slow" even without
  hidden rows.

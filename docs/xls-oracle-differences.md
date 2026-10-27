# XLS: intended differences from SheetJS

`test/xls-oracle.test.ts` compares every committed BIFF8 and BIFF5 fixture
(`test/corpus/generated`, `test/corpus/hidden-rows`, the renamed copy
included) with the SheetJS expectations in `test/corpus/expected` (SheetJS
0.20.3, `cellDates: false`, `cellNF: true`). SheetJS is only a test oracle,
never a dependency. It checks sheet names and visibility, each cell's type
and value (dates as serials; formula cells by their cached result), hidden
rows and columns, merges, the declared dimension, the cell count, and the
defined names (name, scope, formula text, hidden). The hidden-row fixtures,
which make SheetJS's streaming reader loop forever, read with every row,
hidden ones flagged.

`generated/names_formulas.xls` is hand-built for the formula and name
cases: FORMULA records with number, string (STRING), boolean and error
results, a SHRFMLA between a FORMULA and its STRING, a STRING continued by
a CONTINUE, a FORMULA with the BIFF4 record number 0x0406, a self and an
external SUPBOOK, and NAME records (3-D references and areas, sheet spans,
quoted and non-ASCII sheet names, an external reference, deleted and
out-of-range sheets, a deleted-cells reference, whole rows and columns, a
scoped and hidden names, a formula that isn't a reference). The defined-name
differences below are listed name by name in the test, each with SheetJS's
value and ours, so they can't drift.

These are the intended differences:

| case | effect-spreadsheet | SheetJS |
|---|---|---|
| BIFF2, BIFF3, BIFF4 files (Excel 2.x-4.x; `basic.biff2/3/4.xls`), as a bare stream or inside a compound file | `UnsupportedFormat` naming the version (`biff2`, `biff3`, `biff4`) | reads them |
| number with a date/time format | `Date` cell; the test compares it as a serial | number serial, with the format in `z` |
| serial 60 in the 1900 system, negative serials, serials past 9999-12-31, with a date format | stays a `Number` | number with a date `z` |
| formula cells | the cached result only, no `formula` (formulas aren't decompiled) | `f` with the formula text |
| NaN or ±Infinity in NUMBER, RK, MULRK or a FORMULA result | `Error` cell `#NUM!` | `NaN` / `Infinity` numbers |
| -0 | `0` | `-0` |
| an unknown error code in BOOLERR or a FORMULA result | `Error` cell `#N/A` (as XLSX keeps the error text: never a failure) | an error cell with that code |
| a row whose cells come before and after other rows' cells | one row with all its cells, rows in row order | cells by address anyway |
| chart sheets and macro sheets | listed in `sheets`, with no rows | cells made up from the chart's caches |
| sheet dimension | the declared DIMENSIONS record only (absent without one) | `!ref` from the cells it read (differs from DIMENSIONS in some real files) |
| defined names: a reference to an external workbook | only EXTERNSHEET entries of the self SUPBOOK name local sheets; an external one is `[1]Rates!$A$1` (`[n]`: the n-th external SUPBOOK, as XLSX numbers its external links) | `Rates!$A$1` (no workbook) |
| defined names: a deleted sheet, or a sheet index past the workbook's sheets | `#REF!$A$1` | `#REF!$A$1`; a placeholder sheet name when out of range |
| defined names: a reference to deleted cells (PtgRefErr3d, PtgAreaErr3d) | `Sheet!#REF!` (`#REF!` when the sheet is gone too) | `#REF!` |
| defined names: whole columns | `Data!$A:$C`, as Excel writes them (whole rows `Data!$1:$2` agree) | `Data!$A$1:$C$65536` (some real files also give `$A$1:$IV$n` for whole rows) |
| defined names: sheet-name quoting | as Excel: quoted when not name-like or when it reads as a cell reference; letters beyond ASCII need no quotes (`Données!$A$1`); a span `Jan:Mar!$A$1`, quoted as a whole when one name needs it | doesn't quote names that start with a digit or read as a cell reference (`910617!`, `S2!`, `Top10!`, quoted by us); quotes some non-ASCII names (accented Latin `'Données'!`, Cyrillic `'Лист3'!`, Hangul) but not others (kana and CJK ideographs `日本語のみ!`); quotes every span (`'Jan:Mar'!$A$1`) |
| defined names: formulas other than a single 3-D reference or area | `""` (formulas aren't decompiled) | the decompiled formula (e.g. `42`) |
| defined names: `hidden` | the NAME's fHidden flag (e.g. `_xlnm._FilterDatabase`) | never reported for XLS names |
| defined names in BIFF5 files | none | reads them |
| BLANK / MULBLANK cells (styled blanks) | `Empty` cells | no cell (`sheetStubs: false`) |
| one bad record in a sheet (truncated, a shared string out of range, a MULRK whose count doesn't match its columns) | the sheet's rows fail with `MalformedFile` at that record | a value for the cell |
| code pages `TextDecoder` doesn't know (DOS 437, 850, ..., and 28602, which isn't a real ISO-8859 part) for BIFF5 strings | decoded as windows-1252 | its own code-page tables |
| built-in number format 14 | `numFmt: "mm-dd-yy"` (as ECMA-376 lists it) | `z: "m/d/yy"` |
| formatted text (`w`) | not produced | `w` |

BIFF5 fixtures written by SheetJS store non-Latin text as `_`; both readers
give those underscores. BOF versions 0x0500-0x05FF all read as BIFF5
(writers put 0x0500 or another 0x05xx there). Some real files named `.xls` are CSV or XLSX: the
format comes from the bytes (`Spreadsheet.read` never sees a file name), so
they read as what they are.

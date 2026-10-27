# XLSB: intended differences from SheetJS

`test/xlsb-oracle.test.ts` compares every committed XLSB fixture
(`test/corpus/generated`, `test/corpus/hidden-rows`, the renamed copy
included) with the SheetJS expectations in `test/corpus/expected` (SheetJS
0.20.3, `cellDates: false`, `cellNF: true`). SheetJS is only a test oracle,
never a dependency. It checks sheet names and visibility, each cell's type
and value (dates as serials; formula cells by their cached result), hidden
rows and columns, merges, the declared dimension, the cell count, and the
defined names (name, scope, formula text, hidden). The hidden-row fixtures,
which make SheetJS's streaming reader loop forever, read with every row,
hidden ones flagged.

`generated/names_formulas.xlsb` is hand-built (SheetJS's XLSB writer writes
none of these records): every cell record, including rich strings, the four
formula results and the short cells (no column field: the previous cell's
column + 1; SheetJS gives the same cells for the ones in this fixture), date
and custom formats,
hidden rows and columns, merges, hidden and very hidden sheets, a self and
an external book with its external link part, and BrtName records. The
defined-name differences below are listed name by name in the test, each
with SheetJS's value and ours, so they can't drift.

The reader was also run over the XLSB files of the SheetJS test corpus
(not committed: `test/corpus/.cache`, see docs/corpus.md), against the 491
that SheetJS parses within the oracle's limits. Every cell value, hidden row,
hidden column and merge agrees with SheetJS, apart from the differences
below. The 15 corpus files (19 with the committed fixtures) on which
SheetJS's streaming reader never finishes (hidden rows) read in full, every
hidden row flagged (one of them through `Xlsb.read`: it isn't detected as
XLSB, see the table); one other file fails with `MalformedFile` (in the
table too; it is in the corpus twice, as `2013/apachepoi_Simple.xlsb` and
`apache_spreadsheet/xlsb/Simple.xlsb`, the same bytes).

These are the intended differences:

| case | effect-spreadsheet | SheetJS |
|---|---|---|
| number with a date/time format | `Date` cell; the test compares it as a serial | number serial, with the format in `z` |
| serial 60 in the 1900 system, negative serials, serials past 9999-12-31, with a date format | stays a `Number` | number with a date `z` |
| formula cells | the cached result only, no `formula` (formulas are tokens and aren't decompiled) | `f` with the formula text |
| NaN or ±Infinity in a real cell or a formula result | `Error` cell `#NUM!` | `NaN` / `Infinity` numbers |
| -0 | `0` | `-0` |
| an unknown error code | `Error` cell `#N/A` | an error cell with that code |
| BrtCellBlank / BrtShortBlank (styled blanks) | `Empty` cells | stub cells (`t: "z"`), even with `sheetStubs: false` |
| short cell records (BrtShortBlank ... BrtShortIsst, 0x000C-0x0012) | read with no column field: the previous cell's column + 1 in the same row, then a 4-byte style and the value; a short cell that starts a row is in column 0. None occur in the corpus; the synthetic ones in `names_formulas.xlsb` (none of them first in its row) give the same cells as SheetJS | the same cells for those; for a short cell that starts a row, it may carry on from the previous row's last column (not checked: no file has one) |
| record numbers of cells inside future-record blocks (BrtFRTBegin ... BrtFRTEnd), and cell records outside BrtBeginSheetData ... BrtEndSheetData | not cells: ignored (real files reuse 0x0001 inside such blocks). A cell record inside the sheet data before any BrtRowHdr is `MalformedFile` | ignores them too (the corpus files that have them agree cell for cell) |
| chart, dialog and macro sheets | listed in `sheets`, with no rows | cells made up from a chart's caches |
| sheet dimension | the declared BrtWsDim only (absent without one), `A1` for an empty sheet that declares it | `!ref` absent for an empty sheet |
| defined names: a reference to an external workbook | `[1]Rates!$A$1` (`[n]`: the n-th external book, its sheet names from the external link part's BrtSupTabs, as XLSX writes it) | a placeholder (`SH33TJSSRC355!$A$1`) |
| defined names: a sheet index past the workbook's sheets | `#REF!$A$1` | the whole read fails (`empty sheet name`); so no committed fixture has one |
| defined names: the whole sheet (rows 1-1048576, columns A-XFD) | `Sheet1!$1:$1048576`, as Excel writes it (e.g. `2013/apachepoi_3dFormulas.xls.xlsb`) | `Sheet1!$A:$XFD` |
| defined names: a deleted sheet | `#REF!$A$1` | `#REF!$A$1` |
| defined names: a reference to deleted cells (PtgRefErr3d, PtgAreaErr3d) | `Sheet!#REF!` (`#REF!` when the sheet is gone too) | `#REF!` |
| defined names: sheet-name quoting | as Excel, and as for XLS: quoted when not name-like or when it reads as a cell reference (`'S2'!`, `'1'!`, `'Top10'!`); letters beyond ASCII need no quotes; a span `Jan:Mar!$A$1`, quoted as a whole when one name needs it | doesn't quote names that read as a cell reference or start with a digit; quotes non-ASCII names and every span |
| defined names: formulas other than a single 3-D reference or area (unions such as print titles `$A:$B,$1:$2`, constants, functions) | `""` (formulas aren't decompiled) | the decompiled formula |
| defined names: `hidden` | the BrtName's fHidden flag (e.g. `_xlnm._FilterDatabase`) | not reported |
| a broken record (truncated, a string past its record, a shared string out of range, a cell before any row) | `MalformedFile` at that record's offset; e.g. one corpus file's BrtBundleSh has 4 bytes more than MS-XLSB's layout | reads on, sometimes into garbage (that file's sheet names come out as binary) |
| a package whose workbook part isn't `xl/workbook.bin` (one corpus file, `wtf/workbook.bin`) | `Spreadsheet.read` sees a plain `zip` (`UnsupportedFormat`); `Xlsb.read` reads it through the package relationships | reads it |
| built-in number format 14 | `numFmt: "mm-dd-yy"` (as ECMA-376 lists it) | `z: "m/d/yy"` |
| formatted text (`w`) | not produced | `w` |

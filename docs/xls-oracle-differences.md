# XLS: intended differences from SheetJS

`test/xls-oracle.test.ts` compares every committed BIFF8 and BIFF5 fixture
(`test/corpus/generated`, `test/corpus/hidden-rows`, the renamed copy
included) with the SheetJS expectations in `test/corpus/expected` (SheetJS
0.20.3, `cellDates: false`, `cellNF: true`). SheetJS is only a test oracle,
never a dependency. It checks sheet names and visibility, each cell's type
and value (dates as serials), hidden rows and columns, merges, the declared
dimension and the cell count. The hidden-row fixtures, which make SheetJS's
streaming reader loop forever, read with every row, hidden ones flagged.

These are the intended differences:

| case | effect-spreadsheet | SheetJS |
|---|---|---|
| BIFF2, BIFF3, BIFF4 files (Excel 2.x-4.x; `basic.biff2/3/4.xls`) | `UnsupportedFormat` naming the version (`biff2`, `biff3`, `biff4`) | reads them |
| number with a date/time format | `Date` cell; the test compares it as a serial | number serial, with the format in `z` |
| serial 60 in the 1900 system, negative serials, serials past 9999-12-31, with a date format | stays a `Number` | number with a date `z` |
| formula cells | the cached result only, no `formula` (formulas aren't decompiled) | `f` with the formula text |
| defined names | formula text only for a single 3-D reference or area (`Sheet!$A$1:$B$2`), `""` otherwise; none for BIFF5 | the decompiled formula |
| BLANK / MULBLANK cells (styled blanks) | `Empty` cells | no cell (`sheetStubs: false`) |
| one bad record in a sheet (truncated, an unknown error code, a shared string out of range, a MULRK whose count doesn't match its columns) | the sheet's rows fail with `MalformedFile` at that record | a value for the cell |
| code pages `TextDecoder` doesn't know (DOS 437, 850, ...) for BIFF5 strings | decoded as windows-1252 | its own code-page tables |
| built-in number format 14 | `numFmt: "mm-dd-yy"` (as ECMA-376 lists it) | `z: "m/d/yy"` |
| formatted text (`w`) | not produced | `w` |

BIFF5 fixtures written by SheetJS store non-Latin text as `_`; both readers
give those underscores.

# XLSX: intended differences from SheetJS

`test/xlsx-oracle.test.ts` compares every committed xlsx fixture
(`test/corpus/generated`, `test/corpus/hidden-rows`) with the SheetJS
expectations in `test/corpus/expected` (made by `scripts/corpus`, SheetJS
0.20.3, read with `cellDates: false`, `cellNF: true`). SheetJS is only a
test oracle, never a dependency. It checks sheet names and visibility, each
cell's type and value, formulas, hidden rows and the cell count. Where
SheetJS is wrong, `test/corpus/sheetjs-quirks.json` says so and the test
expects the correct value (e.g. `&#x1F363;` truncated to 16 bits, `_x000D_`
dropped).

These are the intended differences:

| case | effect-spreadsheet | SheetJS |
|---|---|---|
| number with a date/time format | `Date` cell (wall clock as UTC, rounded to the millisecond); the test compares it as a serial | number serial, with the format in `z` |
| serial 60 in the 1900 system (the fictional 1900-02-29), negative serials, serials past 9999-12-31, with a date format | stays a `Number` | number with a date `z` |
| `t="d"` cell (ISO date) | `Date` | number serial |
| shared formula, cells after the first | no `formula` (the text is stored only on the first cell; we don't shift references) | the formula expanded for each cell |
| formula without a cached value | `Empty` cell with its `formula` | `t: "z"` stub |
| styled cell without a value | `Empty` cell | no cell (`sheetStubs: false`) |
| built-in number format 14 | `numFmt: "mm-dd-yy"` (the code ECMA-376 lists) | `z: "m/d/yy"` |
| formatted text (`w`) | not produced (PLAN.md: later) | `w` |
| hidden columns, merges, declared dimension | `Workbook.sheetDetails` (not in rows) | `!cols`, `!merges`, `!ref` |

`dates_1904.xlsx`: SheetJS's `sheet_to_json` turns 1904-system serials into
dates with the 1900 epoch (quirk `sheet-to-json-ignores-date1904`); our
`Date`s use the workbook's system.

Not compared: the large SheetJS `test_files` corpus (not committed, see
docs/corpus.md); run the oracle scripts locally for it.

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
| one bad cell (a number, boolean or ISO date that doesn't parse, an unknown `t`, a shared-string index out of range) | the sheet's rows fail with `MalformedFile` at that cell | a value for the cell (and the rest of the sheet) |
| a `<row r>` the same as or below an earlier row's (a row given twice, or rows out of order), also when a row without `r` takes its index from its first cell | `MalformedFile` at that row: rows are streamed in order, each index once, so an earlier row can't take more cells. No file of the SheetJS test corpus has one | merges the cells into one sheet by address |
| a cell whose `r` names another row than its `<row>` | `MalformedFile` at that cell (none in the SheetJS test corpus) | puts the cell at its own address |
| a `<row>` without `r` | the row its first cell's `r` names; without one, the previous row + 1 (a later cell's `r` doesn't decide it, it must agree) | each cell at its own address |
| an inline string (`t="inlineStr"`) with a `<v>` and no `<is>`; an `<is>` in a cell without `t` or `<v>` | the text of the `<v>`; the text of the `<is>` | an empty string; a `t: "z"` stub |
| elements nested in a cell, below its own children (e.g. in its `<extLst>`) | skipped: they don't end the cell or give its value, whatever their names | a nested `<c>` ends the cell and is read as a cell of its own; a nested `<is>` or `<f>` can give the cell's text or formula |
| an element nested in one of its own kind: a `<row>` inside a row, a `<sheetData>` inside the sheet data, an `<si>` inside a shared string, a `<sheets>` inside `<sheets>`, a `<definedName>` inside a defined name (invalid OOXML) | `MalformedFile` at the nested element (in a sheet, the rows completed before it come out first): reading on would drop the open row's cells, the rows or sheets after it, or the open item's text | reads the cells of nested rows at their addresses, but no row after a nested `<sheetData>`; keeps the outer `<si>`'s text before the nested one and reads the next index as an empty string; reads every sheet; gives the inner defined name, twice, with the outer one's end tag in its text |
| a phonetic run (`<rPh>`) inside a `<t>` (invalid OOXML) | only the run is skipped (with the `<t>` elements inside it); the outer `<t>`'s text after the run is kept | the text up to the run's inner `</t>`, the run's markup included; the outer `<t>`'s text after that is lost |
| a defined name whose `localSheetId` names no sheet | no `scope`: it looks workbook-wide | keeps `Sheet: 99` |
| XLM macro sheets (and chart and dialog sheets) | no rows, empty `sheetDetails` | cells, where it reads them |

An impossible ISO date such as `2020-02-30` rolls over to the next valid
day (2020-03-01), as it does in SheetJS.

`dates_1904.xlsx`: SheetJS's `sheet_to_json` turns 1904-system serials into
dates with the 1900 epoch (quirk `sheet-to-json-ignores-date1904`); our
`Date`s use the workbook's system.

Not compared: the large SheetJS `test_files` corpus (not committed, see
docs/corpus.md); run the oracle scripts locally for it.

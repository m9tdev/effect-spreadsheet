# CSV: intended differences from SheetJS

Checked by `test/csv-oracle.test.ts` against SheetJS (the clone in `$SHEETJS_DIR`, default `~/sheetjs`; pinned commit in CI). SheetJS is only a test oracle, never a dependency.

| input | effect-spreadsheet | SheetJS |
|---|---|---|
| `"ab"c` | `abc` | `"ab"c` |
| unterminated `"b…` | `b…` | `"b…` |
| UTF-8 without BOM | decoded as UTF-8 | decoded as Latin-1 |
| `007`, 16+ digit numbers | text | number (precision lost) |
| `true` | `true` (boolean) | `"true"` |
| `50%`, `$5`, `" 1"` | text | 0.5, 5, 1 |
| `2024-01-02` | text | date serial |
| `#N/A` | `Error` cell `"#N/A"` | dropped |
| file starting with a space | normal CSV | PRN path, wrong values |

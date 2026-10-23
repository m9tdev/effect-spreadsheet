# effect-spreadsheet

Effect v4 library for reading spreadsheet files (CSV, XLSX, XLS, XLSB, ...) with its own pure-TypeScript parsers. Work in progress; see PLAN.md. Today only CSV/delimited text reads; the other formats are detected and fail with `UnsupportedFormat`.

It runs in Node, workers and browsers (web-standard APIs only). Every loop has a progress measure, every resource has a limit, and bad input fails with a typed `SpreadsheetError`.

## Usage

```ts
import { Effect, Stream } from "effect"
import { Spreadsheet, SpreadsheetLimits, Workbook } from "effect-spreadsheet"
import * as Csv from "effect-spreadsheet/csv"

// Any file: the format comes from the bytes, never from a file name.
const program = Effect.gen(function*() {
  const workbook = yield* Spreadsheet.read(bytes, { csv: { inferTypes: true } })
  console.log(workbook.format, workbook.sheets.map((s) => s.name))

  // Rows stream lazily, in bounded chunks; the event loop runs between them.
  const records = yield* Workbook.rows(workbook, 0).pipe(
    (rows) => Spreadsheet.toRecords(rows, { header: "first-row" }),
    Stream.runCollect
  )
  return records // [{ name: "ann", age: 31 }, ...]
}).pipe(
  // Tighter limits for this program (per-call `limits` options win over this).
  Effect.provide(SpreadsheetLimits.layer({ maxRows: 100_000 }))
)

// CSV straight from a byte stream (e.g. an upload), any chunking.
const rows = Csv.rows(Stream.fromReadableStream({ evaluate: () => body, onError: (e) => e }), { delimiter: ";" })
```

A `Row` is `{ index, hidden, cells }`; each cell is a tagged object
(`Empty | Text | Number | Boolean | Date | Error`) with its own `column`, plus
`formula` / `numFmt` when the file has them.

Errors: `UnsupportedFormat { format, reason }`, `MalformedFile { format, part?, offset?, reason }`,
`LimitExceeded { limit, max, actual }`, `Encrypted { format }`, `SheetNotFound { sheet }`,
`InvalidOption { option, reason }` (e.g. a bad delimiter or encoding label).

## CSV behaviour

- Encoding: a BOM (UTF-8, UTF-16LE/BE) always wins, even over a forced `encoding`, and is dropped; otherwise `"auto"` decides on the first 64 KiB: UTF-16 by NUL pattern, else UTF-8 if valid, else windows-1252. `encoding` forces any `TextDecoder` label for input without a BOM.
- Delimiter: `delimiter` option, else an Excel `sep=X` first line, else sniffed among `,` `\t` `;` `|` (a count consistent across lines wins, then frequency).
- RFC 4180 quoting; LF / CRLF / CR line breaks; ragged rows; blank lines are empty rows; empty fields produce no cell.
- Lenient by default: text after a closing quote is appended (`"ab"c` → `abc`), an unterminated quote runs to end of input. `strict: true` makes both `MalformedFile`.
- `inferTypes` (default `false`): booleans (any case), plain decimals (no leading zeros, ≤ 15 significant digits), Excel error literals. Never dates, `%`, currency, thousands separators or formulas.

### Intended differences from SheetJS

Checked by `test/csv-oracle.test.ts` against `~/sheetjs`:

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

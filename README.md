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

Bad input and bad reader options are typed errors; bad limits are not. A delimiter that isn't
one usable character or an unknown encoding label fails with `InvalidOption`, but a limit that
isn't a non-negative integer or `Infinity` (e.g. `maxRows: NaN`) is a defect (it dies), because
it would otherwise silently switch the check off.

`Spreadsheet.toRecords` takes its limits (`maxCellsPerSheet` for cells, `maxRecordFields` for
`defval` fills) from the context where the records are produced, plus its own `limits` option,
not from the workbook the rows were read from.

## CSV behaviour

- Encoding: a BOM (UTF-8, UTF-16LE/BE) always wins, even over a forced `encoding`, and is dropped; otherwise `"auto"` decides on the first 64 KiB: UTF-16 by NUL pattern, else UTF-8 if valid, else windows-1252. `encoding` forces any `TextDecoder` label for input without a BOM.
- Delimiter: `delimiter` option, else an Excel `sep=X` first line, else sniffed among `,` `\t` `;` `|` (a count consistent across lines wins, then frequency).
- RFC 4180 quoting; LF / CRLF / CR line breaks; ragged rows; blank lines are empty rows; empty fields produce no cell.
- Lenient by default: text after a closing quote is appended (`"ab"c` → `abc`), an unterminated quote runs to end of input. `strict: true` makes both `MalformedFile`.
- `inferTypes` (default `false`): booleans (any case), decimal numbers, Excel error literals. A number is an optional sign, digits with no leading zeros (`0.5` is fine, `007` stays text), an optional fraction (`1.`, `.5`) and an optional exponent (`1e21`, `1E-3`), with at most 15 digits, not counting leading zeros (`1000000000000000` and `1.000000000000000` stay text), and a finite value (`1e999` stays text); `-0` becomes `0`. Never dates, `%`, currency, thousands separators, hex or formulas.

### Comparison with other readers

The test suite compares CSV results with a widely used JavaScript spreadsheet library as a
test-only oracle. The intended differences are listed in
[docs/csv-oracle-differences.md](https://github.com/m9tdev/effect-spreadsheet/blob/main/docs/csv-oracle-differences.md) (CSV reading) and
[docs/records-oracle-differences.md](https://github.com/m9tdev/effect-spreadsheet/blob/main/docs/records-oracle-differences.md) (`toRecords`).

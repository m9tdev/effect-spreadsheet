# effect-spreadsheet

Effect v4 library for reading spreadsheet files (CSV, XLSX, XLS, XLSB, ...) with its own pure-TypeScript parsers. Work in progress; see PLAN.md. Today CSV/delimited text, XLSX/XLSM and XLS (BIFF8 and BIFF5) read; the other formats are detected and fail with `UnsupportedFormat`.

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

The bytes passed to `read` stay in use after it returns: the workbook keeps views of them, not
copies (XLSX zip entries, a bare BIFF stream, and an XLS workbook stream laid out in consecutive
sectors of its compound file), and sheets are read from them when their rows are. Don't change
those bytes while the workbook is in use; pass a copy if the buffer is reused.

## CSV behaviour

- Encoding: a BOM (UTF-8, UTF-16LE/BE) always wins, even over a forced `encoding`, and is dropped; otherwise `"auto"` decides on the first 64 KiB: UTF-16 by NUL pattern, else UTF-8 if valid, else windows-1252. `encoding` forces any `TextDecoder` label for input without a BOM.
- Delimiter: `delimiter` option, else an Excel `sep=X` first line, else sniffed among `,` `\t` `;` `|` (a count consistent across lines wins, then frequency).
- RFC 4180 quoting; LF / CRLF / CR line breaks; ragged rows; blank lines are empty rows; empty fields produce no cell.
- Lenient by default: text after a closing quote is appended (`"ab"c` → `abc`), an unterminated quote runs to end of input. `strict: true` makes both `MalformedFile`.
- `inferTypes` (default `false`): booleans (any case), decimal numbers, Excel error literals. A number is an optional sign, an integer part without extra leading zeros (`0.5` is fine, `007` stays text), an optional fraction (`1.`, `.5`) and an optional exponent (`1e21`, `1E-3`), with at most 15 digits before the exponent, not counting zeros before the first non-zero digit (`0.000000000000001` is fine; `1000000000000000`, `1.000000000000000` and `1234567890123456e-10` stay text; `123456789012345e10` is a number), and a finite value (`1e999` stays text); `-0` becomes `0`, and so does any value that rounds to 0 as a double (`1e-999` and `-1e-400` become the Number `0`). Values with surrounding whitespace stay text (`" 1"`), and quoted values are inferred too (`"1"` gives the Number `1`). Error literals are case-sensitive and recognised only as `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` and `#GETTING_DATA` (`#n/a` stays text). Never dates, `%`, currency, thousands separators, hex or formulas.

## XLSX behaviour

- `effect-spreadsheet/xlsx` (`Xlsx.read(bytes)`), or `Spreadsheet.read`, which detects XLSX/XLSM from the zip's parts. `.xlsm` reads like `.xlsx`; macros are ignored.
- The workbook part is found through the package relationships. `workbook.sheets` gives names and visibility (`visible`, `hidden`, `veryHidden`); `workbook.definedNames` the defined names (`name`, `formula` text, `scope` sheet index, `hidden`), as metadata only.
- Sheets are read lazily: each `Workbook.rows` run inflates and parses that sheet as a stream, in bounded pieces. Hidden rows are emitted with `hidden: true`. Rows the file leaves out are not emitted (rows are sparse); nothing is sized by the declared dimension.
- Values: shared, inline and formula strings are `Text` (`_xHHHH_` escapes decoded, phonetic runs skipped); booleans, error codes and numbers keep their type. A number whose number format shows a date or time is a `Date` (the wall-clock time as UTC, rounded to the millisecond, 1900 or 1904 date system), as is an ISO-dated cell (`t="d"`: without a zone the wall clock as UTC, with an offset the instant it names). Serials that are no real date (negative, past 9999-12-31, or the 1900 system's fictional 29 February 1900) stay numbers. Cells without a value are `Empty`.
- `numFmt` is the number-format code for numbers and dates (built-in ids as ECMA-376 lists them); `formula` is the formula text as stored (a shared formula only on its first cell). Formulas are not evaluated.
- Zip bombs: all entries of an archive share one inflate budget, `max(inflateFloorBytes, maxInflateRatio × archive size)` (64 MiB and 200 by default), never more than `maxDecompressedBytes` (1 GiB).

## XLS behaviour

- `effect-spreadsheet/xls` (`Xls.read(bytes)`), or `Spreadsheet.read`, which detects an `.xls` from its compound file's workbook stream. BIFF8 (Excel 97-2003) and BIFF5 (Excel 5/95) workbooks read, also as a bare workbook stream. BIFF2, BIFF3 and BIFF4 files (Excel 2.x-4.x) fail with `UnsupportedFormat` naming the version (`biff2`, `biff3`, `biff4`).
- The same `Workbook` API as XLSX: `workbook.sheets` (names, visibility), `workbook.definedNames` (BIFF8; a single 3-D reference or area is written as Excel writes it in XLSX: `Sheet!$A$1`, `'My Sheet'!$A$1:$B$2`, `Jan:Mar!$A$1`, whole rows `$1:$2` and columns `$A:$C`, `[1]Sheet!$A$1` for an external workbook, `#REF!` forms for deleted sheets or cells; other formulas `""`), `Workbook.rows` (hidden rows flagged, never skipped; chart and macro sheets have no rows), `Workbook.sheetDetails` (hidden columns, merges, the declared DIMENSIONS).
- Values: shared strings and labels are `Text` (BIFF5 strings in the workbook's code page via `TextDecoder`; unknown code pages as windows-1252); numbers (NUMBER, RK, MULRK), booleans and error codes keep their type (NaN and infinities are `#NUM!` errors, unknown error codes `#N/A`); formula cells give their cached result (formula text isn't decompiled). Dates and `numFmt` work as in XLSX. Styled blanks are `Empty`.
- A password-protected workbook, or an encrypted OOXML package in a compound file, fails with `Encrypted`.

### Comparison with other readers

The test suite compares CSV results with a widely used JavaScript spreadsheet library as a
test-only oracle. The intended differences are listed in
[docs/csv-oracle-differences.md](https://github.com/m9tdev/effect-spreadsheet/blob/main/docs/csv-oracle-differences.md) (CSV reading) and
[docs/records-oracle-differences.md](https://github.com/m9tdev/effect-spreadsheet/blob/main/docs/records-oracle-differences.md) (`toRecords`) and
[docs/xlsx-oracle-differences.md](https://github.com/m9tdev/effect-spreadsheet/blob/main/docs/xlsx-oracle-differences.md) (XLSX) and
[docs/xls-oracle-differences.md](https://github.com/m9tdev/effect-spreadsheet/blob/main/docs/xls-oracle-differences.md) (XLS).

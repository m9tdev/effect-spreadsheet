# toRecords: intended differences from SheetJS

`Spreadsheet.toRecords` is compared with SheetJS `sheet_to_json` (object
mode, `raw: true`) as a test-only oracle; SheetJS is never a dependency.
These are the intended differences. In the examples the CSV is read with
`inferTypes: true`; SheetJS reads it with its defaults.

Items 3 and 4 are new since the record-key rewrite; items 1 and 2 were
already there before it.

## 1. `__EMPTY` keys are numbered by first use, not by column

A column without a header name gets its key when a data cell first appears
in it, so the numbering follows the data, not the column order.

```
a,,b,,c
1,,3,4,5
,2,,,
```

| | record 1 | record 2 |
|---|---|---|
| effect-spreadsheet | `{ a: 1, b: 3, __EMPTY: 4, c: 5 }` | `{ __EMPTY_1: 2 }` |
| SheetJS | `{ a: 1, b: 3, __EMPTY_1: 4, c: 5 }` | `{ __EMPTY: 2 }` |

## 2. `defval` doesn't fill columns with a blank header

`defval` fills only the columns the header names, so a blank header column
doesn't get a key in every record.

```
a,,b
1
```

with `defval: null`:

| | record |
|---|---|
| effect-spreadsheet | `{ a: 1, b: null }` |
| SheetJS | `{ a: 1, __EMPTY: null, b: null }` |

## 3. Repeated names in an explicit header are made unique (new)

`header: ["x", "x"]` on `1,2`:

| | record |
|---|---|
| effect-spreadsheet | `{ x: 1, x_1: 2 }` |
| SheetJS | `{ x: 2 }` |

Intended: no column is silently lost. The names are made unique the same
way as header-row names.

## 4. An explicit `""` header entry gives `__EMPTY` (new)

`header: ["x", "", "y"]` on `1,2,3`:

| | record |
|---|---|
| effect-spreadsheet | `{ x: 1, __EMPTY: 2, y: 3 }` |
| SheetJS | `{ x: 1, "": 2, y: 3 }` |

An explicit `""` is treated like a blank header cell: it gets an `__EMPTY…`
key when data appears in its column, and no `defval` fill. A hole,
`undefined` or `null` entry drops the column in both.

## Not a difference: boolean header cells

A header cell `TRUE` (a Boolean with `inferTypes`) gives the key `"TRUE"`,
as in SheetJS. A lowercase `true` header still gives `"TRUE"` here and
`"true"` in SheetJS, because SheetJS doesn't infer lowercase booleans in CSV
(see [csv-oracle-differences.md](csv-oracle-differences.md)).

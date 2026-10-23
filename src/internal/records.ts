// Portions derived from SheetJS (bits/90_utils.js), Copyright (C) SheetJS LLC, Apache-2.0.
/**
 * Rows -> records (like SheetJS `sheet_to_json` with object output).
 *
 * @internal
 */
import * as Cell from "../Cell.ts"
import type { Row } from "../Workbook.ts"

export type CellValue = string | number | boolean | Date

export interface RecordsOptions {
  /**
   * - `"first-row"`: the first row is the header.
   * - a number: the row with that `index` is the header; earlier rows are skipped.
   * - an array: column `i` is named `header[i]`; columns past its end are dropped.
   */
  readonly header: "first-row" | number | ReadonlyArray<string>
  /** Value for header columns the row has no cell in. Default: no key. */
  readonly defval?: unknown
  /** Keep rows without any value (as `{}`, or filled with `defval`). Default `false`. */
  readonly blankRows?: boolean | undefined
}

/** Sets a key, also when it is `__proto__` (an own property then, never the prototype). */
const set = (record: Record<string, unknown>, key: string, value: unknown): void => {
  if (key === "__proto__") Object.defineProperty(record, key, { value, enumerable: true, writable: true, configurable: true })
  else record[key] = value
}

const headerText = (cell: Cell.Cell): string => {
  const v = Cell.value(cell)
  return v === undefined ? "" : v instanceof Date ? v.toISOString() : String(v)
}

/**
 * Stateful converter: feed rows in order, get records. Header names are made
 * unique like SheetJS: an empty header is `__EMPTY`, repeats get `_1`, `_2`, ...
 */
export class RecordBuilder {
  private names: Array<string | undefined> | undefined
  private readonly fixed: boolean
  private readonly used = new Map<string, number>()
  private readonly headerRow: number | undefined
  private readonly firstRow: boolean
  private headerDone: boolean

  private readonly options: RecordsOptions

  constructor(options: RecordsOptions) {
    this.options = options
    const header = options.header
    this.fixed = typeof header !== "string" && typeof header !== "number"
    this.firstRow = header === "first-row"
    this.headerRow = typeof header === "number" ? header : undefined
    if (typeof header !== "string" && typeof header !== "number") {
      this.names = [...(header as ReadonlyArray<string>)]
      this.headerDone = true
    } else {
      this.headerDone = false
    }
  }

  private unique(base: string): string {
    const count = this.used.get(base) ?? 0
    if (count === 0) {
      this.used.set(base, 1)
      return base
    }
    let n = count
    let name = `${base}_${n}`
    // progress: n increases; at most (number of names so far) collisions
    while (this.used.has(name)) name = `${base}_${++n}`
    this.used.set(base, n + 1)
    this.used.set(name, 1)
    return name
  }

  private setHeader(row: Row | undefined): void {
    const names: Array<string | undefined> = []
    if (row !== undefined) {
      const last = row.cells.length > 0 ? row.cells[row.cells.length - 1]!.column : -1
      let k = 0
      // progress: c increases up to `last`; k only moves forward through cells
      for (let c = 0; c <= last; c++) {
        const cell = row.cells[k]
        let text = ""
        if (cell !== undefined && cell.column === c) {
          text = headerText(cell)
          k++
        }
        names.push(this.unique(text === "" ? "__EMPTY" : text))
      }
    }
    this.names = names
    this.headerDone = true
  }

  private nameOf(column: number): string | undefined {
    const names = this.names!
    if (column < names.length) return names[column]
    if (this.fixed) return undefined
    // A column past the header row: named lazily, in order of appearance.
    // progress: names grows by one per iteration up to `column`
    while (names.length <= column) names.push(this.unique("__EMPTY"))
    return names[column]
  }

  /** Returns the record for `row`, or undefined if the row is the header, skipped, or blank. */
  push(row: Row): Record<string, unknown> | undefined {
    if (!this.headerDone) {
      if (this.firstRow) {
        this.setHeader(row)
        return undefined
      }
      const at = this.headerRow!
      if (row.index < at) return undefined
      if (row.index === at) {
        this.setHeader(row)
        return undefined
      }
      // The header row isn't in the data (rows can be sparse): no names.
      this.setHeader(undefined)
    }
    const record: Record<string, unknown> = {}
    let empty = true
    const defval = this.options.defval
    if (defval !== undefined) {
      const names = this.names!
      // progress: i increases up to names.length
      for (let i = 0; i < names.length; i++) {
        const name = names[i]
        if (name !== undefined) set(record, name, defval)
      }
    }
    for (const cell of row.cells) {
      const v = Cell.value(cell)
      const name = this.nameOf(cell.column)
      if (name === undefined) continue
      if (v === undefined) {
        if (defval !== undefined) set(record, name, defval)
        continue
      }
      set(record, name, v)
      empty = false
    }
    if (empty && !this.options.blankRows) return undefined
    return record
  }
}

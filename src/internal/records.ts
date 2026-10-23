// Portions derived from SheetJS (bits/90_utils.js), Copyright (C) SheetJS LLC, Apache-2.0.
/**
 * Rows -> records (like SheetJS `sheet_to_json` with object output).
 *
 * @internal
 */
import * as Cell from "../Cell.ts"
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import { LimitExceeded } from "../SpreadsheetError.ts"
import type { LimitOverrides, Limits } from "../SpreadsheetLimits.ts"
import type { Row } from "../Workbook.ts"

export interface RecordsOptions<D = never> {
  /**
   * - `"first-row"`: the first row is the header.
   * - a number: the row with that `index` is the header; earlier rows are skipped.
   * - an array: column `i` is named `header[i]`; columns past its end are dropped.
   */
  readonly header: "first-row" | number | ReadonlyArray<string>
  /** Value for header-named columns the row has no value in. Default: no key. */
  readonly defval?: D
  /** Keep rows without any value (as `{}`, or filled with `defval`). Default `false`. */
  readonly blankRows?: boolean | undefined
  /**
   * Overrides on top of the `SpreadsheetLimits` in the context where the
   * records are produced (when the stream runs, or the Effect runs), not the
   * limits a workbook was read with. Per run, every cell of the rows (Empty
   * included) counts against `maxCellsPerSheet` and every `defval` fill
   * against `maxRecordFields`.
   */
  readonly limits?: LimitOverrides | undefined
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
 *
 * Cost follows the cells present: names live in a sparse map, a column the
 * header row has no cell for is named (`__EMPTY_k`) only when a data cell
 * shows up in it, and `defval` fills only the header-named columns. For the
 * whole run, cells (Empty included) count against `maxCellsPerSheet` and
 * `defval` fills against `maxRecordFields`.
 */
export class RecordBuilder {
  private readonly names = new Map<number, string>()
  /** Header-named columns in column order: where `defval` goes. */
  private headerNames: Array<string> = []
  private readonly fixed: boolean
  private readonly used = new Map<string, number>()
  private readonly headerRow: number | undefined
  private readonly firstRow: boolean
  private headerDone: boolean
  private cells = 0
  private fills = 0
  /** Keys written so far (cells and fills); callers diff it to size batches. */
  keys = 0

  private readonly options: RecordsOptions<unknown>
  private readonly maxCells: number
  private readonly maxFills: number

  constructor(options: RecordsOptions<unknown>, limits: { readonly maxCellsPerSheet: number; readonly maxRecordFields: number }) {
    this.options = options
    this.maxCells = limits.maxCellsPerSheet
    this.maxFills = limits.maxRecordFields
    const header = options.header
    this.firstRow = header === "first-row"
    this.headerRow = typeof header === "number" ? header : undefined
    if (typeof header !== "string" && typeof header !== "number") {
      this.fixed = true
      const given = header as ReadonlyArray<string>
      // progress: i increases up to given.length
      for (let i = 0; i < given.length; i++) {
        const name = given[i]
        if (name === undefined) continue
        this.names.set(i, name)
        this.headerNames.push(name)
      }
      this.headerDone = true
    } else {
      this.fixed = false
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
    if (row !== undefined) {
      // progress: one header cell per iteration (only cells present)
      for (const cell of row.cells) {
        if (this.names.has(cell.column)) continue
        const text = headerText(cell)
        const name = this.unique(text === "" ? "__EMPTY" : text)
        this.names.set(cell.column, name)
        this.headerNames.push(name)
      }
    }
    this.headerDone = true
  }

  private nameOf(column: number): string | undefined {
    const known = this.names.get(column)
    if (known !== undefined || this.fixed) return known
    // A column without a header cell: named when a data cell first appears in it.
    const name = this.unique("__EMPTY")
    this.names.set(column, name)
    return name
  }

  private countCells(n: number): LimitExceeded | undefined {
    this.cells += n
    this.keys += n
    return this.cells > this.maxCells
      ? new LimitExceeded({ limit: "maxCellsPerSheet", max: this.maxCells, actual: this.cells })
      : undefined
  }

  private countFills(n: number): LimitExceeded | undefined {
    this.fills += n
    this.keys += n
    return this.fills > this.maxFills
      ? new LimitExceeded({ limit: "maxRecordFields", max: this.maxFills, actual: this.fills })
      : undefined
  }

  /**
   * Returns the record for `row`; undefined if the row is the header,
   * skipped, or blank; or the limit error.
   */
  push(row: Row): Record<string, unknown> | undefined | LimitExceeded {
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
      const over = this.countFills(this.headerNames.length)
      if (over !== undefined) return over
      // progress: one header name per iteration
      for (const name of this.headerNames) set(record, name, defval)
    }
    const over = this.countCells(row.cells.length)
    if (over !== undefined) return over
    // progress: one cell per iteration
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

/** Keys written per emitted batch: the stream yields between batches. */
export const BATCH_KEYS = 64 * 1024

/**
 * Rows -> records as a stream. Records come out in batches of about
 * {@link BATCH_KEYS} keys with a yield before each, so a wide header times
 * many rows can't block the event loop. When a limit is hit, the records
 * already built are emitted first, then the stream fails.
 */
export const recordsStream = <E, R>(
  rows: Stream.Stream<Row, E, R>,
  options: RecordsOptions<unknown>,
  limits: Limits
): Stream.Stream<Record<string, unknown>, E | LimitExceeded, R> =>
  Stream.transformPull(rows, (upstream) =>
    Effect.sync(() => {
      const builder = new RecordBuilder(options, limits)
      const pullUpstream = Pull.catchDone(upstream, () => Effect.succeed(undefined))
      let chunk: ReadonlyArray<Row> = []
      let i = 0
      let upstreamDone = false
      let failure: LimitExceeded | undefined
      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Record<string, unknown>>, E | LimitExceeded> = Effect.gen(
        function*() {
          // progress: each iteration consumes >= 1 row of `chunk` (i grows),
          // pulls one upstream chunk, or ends (done / failure).
          while (true) {
            if (failure !== undefined) return yield* Effect.fail(failure)
            if (i >= chunk.length) {
              if (upstreamDone) return yield* Cause.done()
              const next = yield* pullUpstream
              if (next === undefined) upstreamDone = true
              else {
                chunk = next
                i = 0
              }
              continue
            }
            yield* Effect.yieldNow
            const out: Array<Record<string, unknown>> = []
            const until = builder.keys + BATCH_KEYS
            // progress: i grows by one per iteration, bounded by chunk.length
            while (i < chunk.length && builder.keys < until) {
              const record = builder.push(chunk[i++]!)
              if (record instanceof LimitExceeded) {
                failure = record
                break
              }
              if (record !== undefined) out.push(record)
            }
            if (out.length > 0) return out as unknown as Arr.NonEmptyReadonlyArray<Record<string, unknown>>
          }
        }
      )
      return pull
    }))

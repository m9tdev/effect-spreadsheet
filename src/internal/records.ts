/**
 * Rows -> records: one object per data row, keyed by header names.
 *
 * @internal
 */
import * as Cell from "../Cell.ts"
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import { LimitExceeded } from "../SpreadsheetError.ts"
import type { LimitOverrides, Limits } from "../SpreadsheetLimits.ts"
import type { Row } from "../Workbook.ts"
import * as RecordKeys from "./recordKeys.ts"

export interface RecordsOptions<D = never> {
  /**
   * - `"first-row"`: the first row is the header.
   * - a number: the row with that `index` is the header; earlier rows are skipped.
   * - an array: column `i` is named `header[i]` (repeats made unique the same
   *   way); columns past its end, and columns whose entry is a hole,
   *   `undefined` or `null`, are dropped. An entry `""` is treated like a
   *   blank header cell: the column gets an `__EMPTY…` key when data appears
   *   in it, and no `defval` fill.
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

/**
 * A header cell as a record-key name: text as is, numbers as `String(v)`,
 * booleans as a spreadsheet shows them (`TRUE`/`FALSE`), a Date as ISO
 * text, an error cell its code; Empty gives no name (the column is then
 * named `__EMPTY…` like any unnamed column).
 */
const headerName = (cell: Cell.Cell): string | undefined => {
  const v = Cell.value(cell)
  if (v === undefined) return undefined
  return v instanceof Date ? v.toISOString() : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v)
}

/**
 * Stateful converter: feed rows in order, get records. Keys come from
 * `recordKeys` (one namer per run): repeated header names get `_1`, `_2`,
 * ...; a column without a header name gets `__EMPTY…` when a data cell first
 * shows up in it. `defval` fills only the header-named columns.
 *
 * For the whole run, cells (Empty included) count against `maxCellsPerSheet`
 * and `defval` fills against `maxRecordFields`.
 */
export class RecordBuilder {
  /** The key namer, once the header is known. */
  private namer: RecordKeys.RecordKeys | undefined
  /** With an explicit `header: [...]`, columns at or past its length are dropped. */
  private readonly explicitLength: number | undefined
  /** Explicit-header columns given no name (a hole, `undefined`, `null`): dropped. */
  private readonly dropped = new Set<number>()
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
    if (typeof header !== "string" && typeof header !== "number") {
      const names = header as ReadonlyArray<string | null | undefined>
      this.explicitLength = names.length
      // An index loop, not `map`: `map` keeps holes, and a hole, `undefined`
      // or `null` drops that column. Strings (also "") are names; explicit
      // names go through the same de-duplication: ["x", "x"] -> x, x_1.
      // progress: column increases by one up to names.length
      const entries: Array<RecordKeys.HeaderEntry> = []
      for (let column = 0; column < names.length; column++) {
        const name = names[column]
        if (typeof name === "string") entries.push([column, name])
        else this.dropped.add(column)
      }
      this.namer = RecordKeys.make(entries)
    } else {
      this.explicitLength = undefined
    }
  }

  /** Builds the namer from the header row's cells (only the cells present). */
  private setHeader(row: Row | undefined): RecordKeys.RecordKeys {
    const entries: Array<RecordKeys.HeaderEntry> = []
    // progress: one header cell per iteration
    if (row !== undefined) for (const cell of row.cells) entries.push([cell.column, headerName(cell)])
    this.namer = RecordKeys.make(entries)
    return this.namer
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
    let namer = this.namer
    if (namer === undefined) {
      const header = this.options.header
      if (header === "first-row") {
        this.setHeader(row)
        return undefined
      }
      const at = header as number
      if (row.index < at) return undefined
      if (row.index === at) {
        this.setHeader(row)
        return undefined
      }
      // The header row isn't in the data (rows can be sparse): no names.
      namer = this.setHeader(undefined)
    }
    const record: Record<string, unknown> = {}
    let empty = true
    const defval = this.options.defval
    if (defval !== undefined) {
      const named = namer.named
      // Only header-named columns without a value in this row are fills;
      // count them before writing anything. progress: one cell per iteration.
      let valued = 0
      for (const cell of row.cells) {
        if (cell._tag !== "Empty" && namer.headerKey(cell.column) !== undefined) valued++
      }
      const over = this.countFills(Math.max(0, named.length - valued))
      if (over !== undefined) return over
      // every header key is written once below (values overwrite): batch by writes
      this.keys += Math.min(valued, named.length)
      // progress: one header-named column per iteration
      for (const [, key] of named) set(record, key, defval)
    }
    const over = this.countCells(row.cells.length)
    if (over !== undefined) return over
    const limit = this.explicitLength
    // progress: one cell per iteration
    for (const cell of row.cells) {
      if (limit !== undefined && (cell.column >= limit || this.dropped.has(cell.column))) continue
      const key = namer.key(cell.column)
      const v = Cell.value(cell)
      if (v === undefined) {
        if (defval !== undefined) set(record, key, defval)
        continue
      }
      set(record, key, v)
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
            // progress: i grows by one per iteration, bounded by chunk.length.
            // Rows that give no record (the header, rows before it, blank
            // rows) add few or no keys, so they don't end the batch; their
            // work is bounded by the upstream chunk, which the reader already
            // keeps to one bounded step (and which yielded before it came).
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

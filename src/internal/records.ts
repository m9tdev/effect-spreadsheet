/**
 * Rows -> records: one object per data row, keyed by header names.
 *
 * @internal
 */
import * as Cell from "../Cell.ts"
import { Cause, Effect, Pull, Stream } from "effect"
import { absurd } from "effect/Function"
import type * as Arr from "effect/Array"
import { InvalidOption, LimitExceeded } from "../SpreadsheetError.ts"
import type { LimitOverrides, Limits } from "../SpreadsheetLimits.ts"
import type { Row } from "../Workbook.ts"
import { type Meter, WorkMeter } from "./meter.ts"
import * as RecordKeys from "./recordKeys.ts"

export interface RecordsOptions<D = never> {
  /**
   * - `"first-row"`: the first row is the header.
   * - a number: the row with that `index` (a safe integer, 0 or more) is the
   *   header; earlier rows are skipped.
   * - an array: column `i` is named `header[i]` (repeats made unique the same
   *   way); columns past its end, and columns whose entry is a hole,
   *   `undefined` or `null`, are dropped. An entry `""` is treated like a
   *   blank header cell: the column gets an `__EMPTY…` key when data appears
   *   in it, and no `defval` fill.
   *
   * Any other value (another string, a fraction, a negative number, `NaN`,
   * an entry that is not a string, `null` or `undefined`) fails with
   * `InvalidOption` before any row is read.
   */
  readonly header: "first-row" | number | ReadonlyArray<string | null | undefined>
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

/** The columns an explicit `header: [...]` keeps: below `length` and not in `dropped`. */
interface ExplicitColumns {
  readonly length: number
  /** Columns given no name (a hole, `undefined`, `null`). */
  readonly dropped: ReadonlySet<number>
}

/** The header is known: records are named by `namer`. */
interface Ready {
  readonly _tag: "Ready"
  readonly namer: RecordKeys.RecordKeys
  /** Set only for an explicit `header: [...]`. */
  readonly explicit: ExplicitColumns | undefined
}

/**
 * Where a builder is in its run. `header: "first-row"` starts in
 * `AwaitFirstRow`, a row index in `AwaitRow`, an explicit array in `Ready`.
 * Both waiting phases move to `Ready` once and stay there.
 */
type Phase =
  | { readonly _tag: "AwaitFirstRow" }
  | { readonly _tag: "AwaitRow"; readonly at: number }
  | Ready

const invalidHeader = (reason: string): InvalidOption => new InvalidOption({ option: "header", reason })

/**
 * Checks `options.header` at runtime, since the type can be bypassed (plain
 * JavaScript, casts): only the values {@link RecordsOptions.header} documents
 * pass.
 */
export const validateHeader = (header: unknown): InvalidOption | undefined => {
  if (header === "first-row") return undefined
  if (typeof header === "number") {
    return Number.isSafeInteger(header) && header >= 0
      ? undefined
      : invalidHeader(`${header} is not a row index (a safe integer, 0 or more)`)
  }
  if (typeof header === "string") {
    return invalidHeader(`${JSON.stringify(header)} must be "first-row", a row index or an array of names`)
  }
  if (!Array.isArray(header)) {
    const got = header === null ? "null" : typeof header
    return invalidHeader(`must be "first-row", a row index or an array of names, not ${got}`)
  }
  const names: ReadonlyArray<unknown> = header
  // An index loop, so holes are seen as undefined. progress: column
  // increases by one up to names.length
  for (let column = 0; column < names.length; column++) {
    const name = names[column]
    if (typeof name !== "string" && name !== null && name !== undefined) {
      return invalidHeader(`entry ${column} must be a string, null or undefined, not ${typeof name}`)
    }
  }
  return undefined
}

/** The phase a run starts in, decided once from a valid `options.header`. */
const initial = (header: RecordsOptions<unknown>["header"], meter: Meter): Phase => {
  if (header === "first-row") return { _tag: "AwaitFirstRow" }
  if (typeof header === "number") return { _tag: "AwaitRow", at: header }
  const names = header
  // An index loop, not `map`: `map` keeps holes, and a hole, `undefined`
  // or `null` drops that column. Strings (also "") are names; explicit
  // names go through the same de-duplication: ["x", "x"] -> x, x_1.
  // progress: column increases by one up to names.length
  const entries: Array<RecordKeys.HeaderEntry> = []
  const dropped = new Set<number>()
  for (let column = 0; column < names.length; column++) {
    const name = names[column]
    if (typeof name === "string") entries.push([column, name])
    else dropped.add(column)
  }
  return { _tag: "Ready", namer: RecordKeys.make(entries, meter), explicit: { length: names.length, dropped } }
}

/** `Ready` with names from the header row's cells (only the cells present), or none without a row. */
const ready = (row: Row | undefined, meter: Meter): Ready => {
  const entries: Array<RecordKeys.HeaderEntry> = []
  // progress: one header cell per iteration
  if (row !== undefined) for (const cell of row.cells) entries.push([cell.column, headerName(cell)])
  return { _tag: "Ready", namer: RecordKeys.make(entries, meter), explicit: undefined }
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
  private phase: Phase
  private cells = 0
  private fills = 0

  private readonly options: RecordsOptions<unknown>
  private readonly maxCells: number
  private readonly maxFills: number
  /** Where the namers count their key probes. */
  private readonly meter: Meter

  constructor(
    options: RecordsOptions<unknown>,
    limits: { readonly maxCellsPerSheet: number; readonly maxRecordFields: number },
    meter: Meter
  ) {
    this.options = options
    this.maxCells = limits.maxCellsPerSheet
    this.maxFills = limits.maxRecordFields
    this.meter = meter
    this.phase = initial(options.header, meter)
  }

  private countCells(n: number): LimitExceeded | undefined {
    this.cells += n
    return this.cells > this.maxCells
      ? new LimitExceeded({ limit: "maxCellsPerSheet", max: this.maxCells, actual: this.cells })
      : undefined
  }

  private countFills(n: number): LimitExceeded | undefined {
    this.fills += n
    return this.fills > this.maxFills
      ? new LimitExceeded({ limit: "maxRecordFields", max: this.maxFills, actual: this.fills })
      : undefined
  }

  /**
   * Takes `row`: its record goes to `out`, unless the row is the header,
   * skipped, or blank. Returns the keys it wrote, which callers size
   * batches by (see `record`), or the limit error.
   */
  push(row: Row, out: Array<Record<string, unknown>>): number | LimitExceeded {
    const phase = this.phase
    switch (phase._tag) {
      case "Ready":
        return this.record(row, phase, out)
      case "AwaitFirstRow":
        this.phase = ready(row, this.meter)
        return 0
      case "AwaitRow": {
        if (row.index < phase.at) return 0
        if (row.index === phase.at) {
          this.phase = ready(row, this.meter)
          return 0
        }
        // The header row isn't in the data (rows can be sparse): no names,
        // and this row is data.
        const next = ready(undefined, this.meter)
        this.phase = next
        return this.record(row, next, out)
      }
      default:
        return absurd(phase)
    }
  }

  /**
   * The record for a data row, once the header is known, into `out` unless
   * it is blank. Returns its work in keys, to size batches by: one per cell
   * of the row (also a cell that writes none: Empty, repeated or dropped)
   * and, with `defval`, one per header-named column (each written once,
   * as a fill or a value).
   */
  private record(row: Row, phase: Ready, out: Array<Record<string, unknown>>): number | LimitExceeded {
    const namer = phase.namer
    const explicit = phase.explicit
    // One cell per column: the last one, also when it is Empty.
    const cells = lastPerColumn(row.cells)
    const record: Record<string, unknown> = {}
    let empty = true
    const defval = this.options.defval
    if (defval !== undefined) {
      const named = namer.named
      // Only header-named columns without a value in this row are fills;
      // count them before writing anything. progress: one cell per iteration.
      let valued = 0
      for (const cell of cells) {
        if (cell._tag !== "Empty" && namer.headerKey(cell.column) !== undefined) valued++
      }
      const over = this.countFills(named.length - valued)
      if (over !== undefined) return over
      // progress: one header-named column per iteration
      for (const [, key] of named) set(record, key, defval)
    }
    const over = this.countCells(row.cells.length)
    if (over !== undefined) return over
    // progress: one cell per iteration
    for (const cell of cells) {
      if (explicit !== undefined && (cell.column >= explicit.length || explicit.dropped.has(cell.column))) continue
      const v = Cell.value(cell)
      // An Empty cell is not data: a header-named column already holds its
      // defval fill, and an unnamed column gets no key until data appears.
      if (v === undefined) continue
      set(record, namer.key(cell.column), v)
      empty = false
    }
    if (!empty || this.options.blankRows) out.push(record)
    return row.cells.length + (defval === undefined ? 0 : namer.named.length)
  }
}

/**
 * A row's cells with one cell per column: the last cell of a repeated
 * column, also when it is Empty, as the readers keep it. A column stays at
 * the place of its first cell. Rows whose columns strictly ascend (all rows
 * the readers make) are returned as they are.
 */
const lastPerColumn = (cells: ReadonlyArray<Cell.Cell>): ReadonlyArray<Cell.Cell> => {
  let previous = -1
  let ascending = true
  // progress: one cell per iteration
  for (const cell of cells) {
    if (!(cell.column > previous)) {
      ascending = false
      break
    }
    previous = cell.column
  }
  if (ascending) return cells
  // Map.set on a column already present keeps its place and replaces its cell.
  const byColumn = new Map<number, Cell.Cell>()
  // progress: one cell per iteration
  for (const cell of cells) byColumn.set(cell.column, cell)
  return Array.from(byColumn.values())
}

/** Keys written per emitted batch: the stream yields between batches. */
export const BATCH_KEYS = 64 * 1024

/**
 * Rows -> records as a stream. Records come out in batches of about
 * {@link BATCH_KEYS} keys with a yield before each, so a wide header times
 * many rows can't block the event loop. When a limit is hit, the records
 * already built are emitted first, then the stream fails. An invalid
 * `header` fails the stream before `rows` is pulled.
 */
export const recordsStream = <E, R>(
  rows: Stream.Stream<Row, E, R>,
  options: RecordsOptions<unknown>,
  limits: Limits
): Stream.Stream<Record<string, unknown>, E | LimitExceeded | InvalidOption, R> => {
  const invalid = validateHeader(options.header)
  if (invalid !== undefined) return Stream.fail(invalid)
  return Stream.transformPull(rows, (upstream) =>
    Effect.map(WorkMeter, (meter) => {
      const builder = new RecordBuilder(options, limits, meter)
      const pullUpstream = Pull.catchDone(upstream, () => Effect.undefined)
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
            let keys = 0
            // progress: i grows by one per iteration, bounded by chunk.length.
            // Rows that give no record (the header, rows before it, blank
            // rows) add few or no keys, so they don't end the batch; their
            // work is bounded by the upstream chunk, which the reader already
            // keeps to one bounded step (and which yielded before it came).
            while (i < chunk.length && keys < BATCH_KEYS) {
              const written = builder.push(chunk[i++]!, out)
              if (written instanceof LimitExceeded) {
                failure = written
                break
              }
              keys += written
            }
            if (out.length > 0) return out as unknown as Arr.NonEmptyReadonlyArray<Record<string, unknown>>
          }
        }
      )
      return pull
    }))
}

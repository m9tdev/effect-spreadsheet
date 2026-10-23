import { Cause, Effect, Exit, Option, Stream } from "effect"
import { isSpreadsheetError } from "../../src/SpreadsheetError.ts"
import * as Csv from "../../src/csv/index.ts"
import type { SpreadsheetError } from "../../src/SpreadsheetError.ts"
import type { Row } from "../../src/Workbook.ts"
import * as Cell from "../../src/Cell.ts"

export const enc = (s: string): Uint8Array => new TextEncoder().encode(s)

/** Parses and collects all rows (or fails). */
export const csvRows = (
  input: Uint8Array | string | Stream.Stream<Uint8Array>,
  options?: Csv.ReadOptions
): Effect.Effect<Array<Row>, SpreadsheetError> =>
  Stream.runCollect(Csv.rows(typeof input === "string" ? enc(input) : input as Uint8Array, options)).pipe(
    Effect.map((rows) => Array.from(rows))
  )

export const csvExit = (input: Uint8Array | string | Stream.Stream<Uint8Array>, options?: Csv.ReadOptions) =>
  Effect.runPromise(Effect.exit(csvRows(input, options)))

/** Rows as arrays of values (dense up to the last cell, holes as undefined). */
export const toAoa = (rows: ReadonlyArray<Row>): Array<Array<unknown>> =>
  rows.map((row) => {
    const out: Array<unknown> = []
    for (const cell of row.cells) out[cell.column] = Cell.value(cell)
    return Array.from(out)
  })

export const csvAoa = async (input: Uint8Array | string | Stream.Stream<Uint8Array>, options?: Csv.ReadOptions) =>
  toAoa(await Effect.runPromise(csvRows(input, options)))

/** Splits bytes at the given cut points (sorted, deduplicated, in range). */
export const splitAt = (bytes: Uint8Array, cuts: ReadonlyArray<number>): Array<Uint8Array> => {
  const points = [...new Set(cuts.map((c) => Math.max(0, Math.min(bytes.length, c))))].sort((a, b) => a - b)
  const out: Array<Uint8Array> = []
  let last = 0
  for (const p of points) {
    out.push(bytes.subarray(last, p))
    last = p
  }
  out.push(bytes.subarray(last))
  return out
}

export const failureOf = <A, E>(exit: Exit.Exit<A, E>): E | undefined =>
  Exit.isFailure(exit) ? Option.getOrUndefined(Exit.findErrorOption(exit)) : undefined

/**
 * "value" for success, "typed" for a SpreadsheetError failure, "defect" for
 * anything else (die, interrupt, untyped failure) with the pretty cause.
 */
export const classify = <A, E>(exit: Exit.Exit<A, E>): { kind: "value" | "typed" | "defect"; detail?: string } => {
  if (Exit.isSuccess(exit)) return { kind: "value" }
  const cause = exit.cause
  if (Cause.hasDies(cause) || Cause.hasInterrupts(cause)) return { kind: "defect", detail: Cause.pretty(cause) }
  const error = Option.getOrUndefined(Exit.findErrorOption(exit))
  return isSpreadsheetError(error) ? { kind: "typed" } : { kind: "defect", detail: Cause.pretty(cause) }
}

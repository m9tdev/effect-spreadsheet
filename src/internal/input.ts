/**
 * What every reader checks before it looks at the bytes, done once at the
 * entry point: the limits for the call are resolved (the context's, with
 * the call's overrides on top) and validated, and the input size is checked
 * against `maxInputBytes`. Readers get the result as values and don't
 * redo either.
 */
import { Effect } from "effect"
import { LimitExceeded } from "../SpreadsheetError.ts"
import {
  defaults,
  isValidLimit,
  type LimitOverrides,
  type Limits,
  merge,
  SpreadsheetLimits,
  validate
} from "../SpreadsheetLimits.ts"

declare const ValidBrand: unique symbol

/**
 * Limits known to be valid. Made only here, by {@link resolveLimits} and
 * {@link mergeValid} (or {@link mergeLimits}), which validate what they are
 * given. The brand is a type only: nothing is recorded at run time, except
 * for the limits a reader keeps in its workbook ({@link stored}).
 */
export type ValidLimits = Limits & { readonly [ValidBrand]: true }

/**
 * The limits readers keep as `workbook.limits`. `Workbook.make` is public
 * and takes plain `Limits`, so a workbook's limits are only known to be
 * valid when they are one of these objects. Only those are recorded, once
 * per read: detection and the per-call merges of `Workbook.rows` don't
 * pay for it.
 */
const kept = new WeakSet<Limits>()

/** Records `limits` as a workbook's own: for the reader that builds the workbook. */
export const stored = (limits: ValidLimits): ValidLimits => {
  kept.add(limits)
  return limits
}

/** Whether `limits` are the ones a reader kept in its workbook (so valid). */
export const isStored = (limits: Limits): limits is ValidLimits => kept.has(limits)

/**
 * `base` with the defined `overrides` on top. Only the overrides are
 * validated, since `base` already is; an invalid one throws the same
 * `TypeError` as `SpreadsheetLimits.merge` (the first invalid limit in key
 * order, which can only be an override).
 */
export const mergeValid = (base: ValidLimits, overrides: LimitOverrides | undefined): ValidLimits => {
  if (overrides === undefined) return base
  const out: { -readonly [K in keyof Limits]: number } = { ...base }
  let invalid = false
  for (const key of Object.keys(defaults) as Array<keyof Limits>) {
    const value = overrides[key]
    if (value !== undefined) {
      out[key] = value
      if (!isValidLimit(value)) invalid = true
    }
  }
  if (invalid) validate(out) // throws, naming the first invalid limit
  return out as ValidLimits
}

/**
 * `base` with the defined `overrides` on top, as `SpreadsheetLimits.merge`:
 * only the overrides are validated when `base` is a workbook's stored
 * limits, everything otherwise.
 */
export const mergeLimits = (base: Limits, overrides: LimitOverrides | undefined): ValidLimits =>
  isStored(base) ? mergeValid(base, overrides) : merge(base, overrides) as ValidLimits

/**
 * The limits for one call: the context's limits with the call's overrides
 * on top, all validated. Invalid values (in the context or the overrides)
 * make it die, as `SpreadsheetLimits.resolve` does.
 */
export const resolveLimits = (overrides?: LimitOverrides): Effect.Effect<ValidLimits> =>
  Effect.flatMap(SpreadsheetLimits, (limits) => Effect.sync(() => merge(limits, overrides) as ValidLimits))

/** An input whose size has been checked against the limits it is read with. */
export interface Accepted {
  readonly bytes: Uint8Array
  readonly limits: ValidLimits
}

/**
 * Resolves the limits for one call (the context's, with `overrides` on top,
 * all validated) and checks the input size against `maxInputBytes`, once,
 * before any reading; then continues with `read`. Invalid limits die, as
 * with {@link resolveLimits}. `read` is a continuation rather than a
 * `flatMap` on the result so that small calls like `detectFormat` cost no
 * more than the check they replace.
 */
export const accept = <A, E, R>(
  bytes: Uint8Array,
  overrides: LimitOverrides | undefined,
  read: (input: Accepted) => Effect.Effect<A, E, R>
): Effect.Effect<A, E | LimitExceeded, R> =>
  Effect.flatMap(SpreadsheetLimits, (context) =>
    // suspend: invalid limits (merge throws) die when the effect runs
    Effect.suspend((): Effect.Effect<A, E | LimitExceeded, R> => {
      const limits = merge(context, overrides) as ValidLimits
      return bytes.length > limits.maxInputBytes
        ? Effect.fail(new LimitExceeded({ limit: "maxInputBytes", max: limits.maxInputBytes, actual: bytes.length }))
        : read({ bytes, limits })
    }))

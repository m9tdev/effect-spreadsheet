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
 * given.
 */
export type ValidLimits = Limits & { readonly [ValidBrand]: true }

/**
 * Every `ValidLimits` made here. `Workbook.make` is public and takes plain
 * `Limits`, so a workbook's limits are only known to be valid when they are
 * one of these objects.
 */
const made = new WeakSet<Limits>()

const brand = (limits: Limits): ValidLimits => {
  made.add(limits)
  return limits as ValidLimits
}

/** Whether `limits` were made (validated) here. */
export const isValid = (limits: Limits): limits is ValidLimits => made.has(limits)

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
  return brand(out)
}

/**
 * `base` with the defined `overrides` on top: as `SpreadsheetLimits.merge`,
 * validating only the overrides when `base` is already valid.
 */
export const mergeLimits = (base: Limits, overrides: LimitOverrides | undefined): ValidLimits =>
  isValid(base) ? mergeValid(base, overrides) : brand(merge(base, overrides))

/**
 * The limits for one call: the context's limits with the call's overrides
 * on top, all validated. Invalid values (in the context or the overrides)
 * make it die, as `SpreadsheetLimits.resolve` does.
 */
export const resolveLimits = (overrides?: LimitOverrides): Effect.Effect<ValidLimits> =>
  Effect.flatMap(SpreadsheetLimits, (limits) => Effect.sync(() => mergeLimits(limits, overrides)))

/** An input whose size has been checked against the limits it is read with. */
export interface Accepted {
  readonly bytes: Uint8Array
  readonly limits: ValidLimits
}

/**
 * Resolves the limits for one call ({@link resolveLimits}) and checks the
 * input size against `maxInputBytes`, once, before any reading.
 */
export const accept = (bytes: Uint8Array, overrides?: LimitOverrides): Effect.Effect<Accepted, LimitExceeded> =>
  Effect.flatMap(resolveLimits(overrides), (limits) =>
    bytes.length > limits.maxInputBytes
      ? Effect.fail(new LimitExceeded({ limit: "maxInputBytes", max: limits.maxInputBytes, actual: bytes.length }))
      : Effect.succeed({ bytes, limits }))

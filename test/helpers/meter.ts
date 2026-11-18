import { Effect } from "effect"
import { type Meter, type WorkUnit, WorkMeter } from "../../src/internal/meter.ts"

/** A work meter that sums the work reported to it in one unit, ignoring the others. */
export interface Counter extends Meter {
  readonly total: () => number
}

export const counter = (unit: WorkUnit): Counter => {
  let total = 0
  return {
    add: (u, count) => {
      if (u === unit) total += count
    },
    total: () => total
  }
}

/** `effect` with `meter` as its work meter. */
export const metered = <A, E, R>(meter: Meter, effect: Effect.Effect<A, E, R>): Effect.Effect<A, E, R> =>
  Effect.provideService(effect, WorkMeter, meter)

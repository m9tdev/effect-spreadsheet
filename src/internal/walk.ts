/**
 * Long synchronous walks over an input, written as generators and driven in
 * bounded turns.
 *
 * A walk does its work in small steps and counts the cost of each one on
 * its `Turn` (in whatever measure the walk picked: records, bytes, a mix).
 * Once a turn's budget is spent, the walk yields, and the driver gives the
 * event loop a turn before it resumes the walk with a new budget. So a
 * walk keeps its state in its own locals and control flow, rather than in
 * a state machine that can stop and resume anywhere, and still never runs
 * for much more than one turn's budget at a time.
 *
 * @internal
 */
import { Effect, type Result } from "effect"
import { WorkMeter, type WorkUnit } from "./meter.ts"
import { yieldNow } from "./yield.ts"

/** The budget of a walk's current turn (see `drive`). */
export class Turn {
  /** The budget per turn. */
  readonly size: number
  /** What is left of the current turn's budget: at most `size`; 0 or less once it is spent. */
  left: number

  constructor(size: number) {
    this.size = size
    this.left = size
  }

  /**
   * Counts a step's cost: at least 1, whatever it is (so a turn holds at
   * most `size` steps). True once the turn's budget is spent: the walk
   * yields then, before its next step.
   */
  spend(cost: number): boolean {
    // (written so that NaN counts as 1 too)
    this.left -= cost >= 1 ? cost : 1
    return this.left <= 0
  }
}

/**
 * A walk: a generator that yields when its turn's budget is spent (see
 * `Turn.spend`), and returns its value or a typed error.
 */
export type Walk<A, E> = Generator<void, Result.Result<A, E>, void>

/**
 * Runs a walk to its end. `make` gets the walk's `Turn`, of `perTurn`;
 * whenever the walk yields, the driver gives the event loop a turn and
 * resumes it with a fresh budget. Each turn counts one `unit` on the work
 * meter. Interruption takes effect at the yields; the walk is then dropped
 * where it stands.
 */
export const drive = <A, E>(make: (turn: Turn) => Walk<A, E>, perTurn: number, unit: WorkUnit): Effect.Effect<A, E> =>
  Effect.gen(function*() {
    const meter = yield* WorkMeter
    const turn = new Turn(perTurn)
    const walk = make(turn)
    meter.add(unit, 1)
    // progress: the walk's own (each step advances it, and a turn holds at
    // least one step); one turn per iteration
    for (;;) {
      const step = walk.next()
      if (step.done === true) return yield* Effect.fromResult(step.value)
      yield* yieldNow
      meter.add(unit, 1)
      turn.left = perTurn
    }
  })

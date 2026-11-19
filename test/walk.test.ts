/** The driver of generator walks: turn budgets, event-loop turns, the work meter, failures and interruption. */
import { Effect, Exit, Fiber, Result } from "effect"
import { describe, expect, it } from "vitest"
import { drive, Turn, type Walk } from "../src/internal/walk.ts"
import { counter, metered } from "./helpers/meter.ts"
import { countYields } from "./helpers/yields.ts"

/** A walk of `steps` steps of `cost` each, recording the budget left before each step; returns the steps taken. */
const stepping = (steps: number, cost: number, seen: Array<number> = []) => function*(turn: Turn): Walk<number, never> {
  // progress: k up to steps
  for (let k = 0; k < steps; k++) {
    seen.push(turn.left)
    if (turn.spend(cost)) yield
  }
  return Result.succeed(steps)
}

describe("Turn", () => {
  it("counts each step's cost and says when the budget is spent", () => {
    const turn = new Turn(10)
    expect(turn.size).toBe(10)
    expect(turn.spend(4)).toBe(false)
    expect(turn.left).toBe(6)
    expect(turn.spend(5)).toBe(false)
    expect(turn.spend(1)).toBe(true)
    expect(turn.left).toBe(0)
    const over = new Turn(10)
    expect(over.spend(25)).toBe(true)
    expect(over.left).toBe(-15)
  })

  it("counts a step as 1 at least, whatever its cost (zero, negative, fractional, NaN)", () => {
    const turn = new Turn(5)
    expect(turn.spend(0)).toBe(false)
    expect(turn.spend(-100)).toBe(false)
    expect(turn.spend(0.25)).toBe(false)
    expect(turn.spend(Number.NaN)).toBe(false)
    expect(turn.left).toBe(1)
    expect(turn.spend(0)).toBe(true)
    // a cost above 1 counts as is
    const exact = new Turn(5)
    exact.spend(1.5)
    expect(exact.left).toBe(3.5)
  })
})

describe("drive", () => {
  it("runs a walk to its value, with an event-loop turn each time its budget is spent, and a fresh budget after it", async () => {
    const seen: Array<number> = []
    const batches = counter("batches")
    const run = await countYields(() => Effect.runPromise(metered(batches, drive(stepping(1000, 3, seen), 30, "batches"))))
    expect(run.value).toBe(1000)
    // 10 steps per turn: a yield after every 10th step (the last one too), 101 turns
    expect(batches.total()).toBe(101)
    expect(run.yields).toBeGreaterThanOrEqual(100)
    expect(seen.slice(0, 12)).toEqual([30, 27, 24, 21, 18, 15, 12, 9, 6, 3, 30, 27])
  })

  it("starts each turn with the whole budget, even after a step that overspent the last one", async () => {
    const seen: Array<number> = []
    const costs = [25, 1, 1]
    const walk = function*(turn: Turn): Walk<number, never> {
      // progress: one cost per iteration
      for (const cost of costs) {
        seen.push(turn.left)
        if (turn.spend(cost)) yield
      }
      return Result.succeed(costs.length)
    }
    expect(await Effect.runPromise(drive(walk, 10, "batches"))).toBe(3)
    expect(seen).toEqual([10, 10, 9])
  })

  it("counts the work on the meter it finds, in the unit it is given, one per turn", async () => {
    const added: Array<[string, number]> = []
    await Effect.runPromise(metered({ add: (unit, count) => added.push([unit, count]) }, drive(stepping(25, 1, []), 10, "steps")))
    expect(added).toEqual([["steps", 1], ["steps", 1], ["steps", 1]])
  })

  it("doesn't yield for a walk that fits in one turn", async () => {
    const batches = counter("batches")
    const run = await countYields(() => Effect.runPromise(metered(batches, drive(stepping(9, 1, []), 10, "batches"))), batches.total)
    expect(run.value).toBe(9)
    expect(batches.total()).toBe(1)
  })

  it("fails with the walk's error, after the turns it took", async () => {
    const batches = counter("batches")
    const failing = function*(turn: Turn): Walk<number, string> {
      // progress: k up to 50
      for (let k = 0; k < 50; k++) if (turn.spend(1)) yield
      return Result.fail("broken")
    }
    const exit = await Effect.runPromise(Effect.exit(metered(batches, drive(failing, 10, "batches"))))
    expect(exit).toEqual(Exit.fail("broken"))
    // yields after steps 10, 20, 30, 40 and 50
    expect(batches.total()).toBe(6)
  })

  it("is interrupted between turns: the walk takes no step after it", async () => {
    let steps = 0
    const endless = function*(turn: Turn): Walk<never, never> {
      // progress: none on purpose; the test interrupts it
      for (;;) {
        steps++
        if (turn.spend(1)) yield
      }
    }
    const exit = await Effect.runPromise(Effect.gen(function*() {
      const fiber = yield* Effect.forkChild(drive(endless, 100, "batches"))
      // progress: i up to 5
      for (let i = 0; i < 5; i++) yield* Effect.promise(() => new Promise((resolve) => setImmediate(resolve)))
      return yield* Fiber.interrupt(fiber).pipe(Effect.andThen(Fiber.await(fiber)))
    }))
    expect(Exit.hasInterrupts(exit)).toBe(true)
    const at = steps
    expect(at).toBeGreaterThanOrEqual(100)
    await new Promise((resolve) => setTimeout(resolve, 20))
    expect(steps).toBe(at)
  })
})

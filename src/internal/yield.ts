/**
 * The yield the readers take between two bounded steps of work (one per
 * byte slice, record batch or row chunk): it gives the event loop a turn.
 *
 * `Effect.yieldNow` resumes through the fiber's scheduler. Effect's default
 * scheduler takes `setImmediate` when the runtime has one (Node, Bun), else
 * `setTimeout(fn, 0)`, and browsers clamp nested timeouts to at least 4 ms.
 * A read yields hundreds of times, often after under a millisecond of work,
 * so in a browser it would spend most of its time waiting on the clamp. So
 * when the fiber runs on that default scheduler with no `setImmediate`, the
 * yield resumes on a `MessageChannel` message instead: still a macrotask
 * (input, rendering and other tasks get their turn), but not clamped. A
 * scheduler the application provided itself is always used as is.
 *
 * @internal
 */
import { Effect } from "effect"
import * as Scheduler from "effect/Scheduler"

/** Runs a function on a later macrotask. */
export type PostTask = (task: () => void) => void

/**
 * Posts tasks as `MessageChannel` messages, in order. The channel is made
 * on the first post.
 */
export const messageChannelTask = (): PostTask => {
  let port: MessagePort | undefined
  const queue: Array<() => void> = []
  return (task) => {
    if (port === undefined) {
      const channel = new MessageChannel()
      // one message per queued task, so each message runs exactly one
      channel.port1.onmessage = () => queue.shift()?.()
      port = channel.port2
    }
    queue.push(task)
    port.postMessage(undefined)
  }
}

/** What Effect's default scheduler resumes fibers with: `setImmediate`, or `setTimeout(0)` without it. */
const DEFAULT_SET_IMMEDIATE = new Scheduler.MixedScheduler().setImmediate

/**
 * Whether the scheduler is Effect's default one: a `MixedScheduler` that
 * resumes through Effect's own fallback. Anything else (a synchronous one,
 * one given its own `setImmediate`, another implementation) was chosen by
 * the application.
 */
export const isDefaultScheduler = (scheduler: Scheduler.Scheduler): boolean =>
  scheduler instanceof Scheduler.MixedScheduler && scheduler.setImmediate === DEFAULT_SET_IMMEDIATE

/**
 * A yield that resumes through `post` when the fiber's scheduler is the
 * default one, and is `Effect.yieldNow` otherwise. Interruption while it
 * waits is immediate; the posted task then finds the fiber resumed and
 * does nothing.
 */
export const yieldThrough = (post: PostTask): Effect.Effect<void> => {
  const posted = Effect.callback<void>((resume) => {
    post(() => resume(Effect.void))
  })
  return Effect.withFiber((fiber) => isDefaultScheduler(fiber.getRef(Scheduler.Scheduler)) ? posted : Effect.yieldNow)
}

/**
 * Picks the yield for a runtime: `Effect.yieldNow` where Effect already
 * resumes through `setImmediate`, or where there is no `MessageChannel`;
 * else a yield through a `MessageChannel`.
 */
export const yieldFor = (runtime: { readonly setImmediate?: unknown; readonly MessageChannel?: unknown }): Effect.Effect<void> =>
  typeof runtime.setImmediate === "function" || typeof runtime.MessageChannel !== "function"
    ? Effect.yieldNow
    : yieldThrough(messageChannelTask())

/** The yield every reader uses (instead of `Effect.yieldNow`). */
export const yieldNow: Effect.Effect<void> = yieldFor(globalThis)

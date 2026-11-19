import { Effect, Exit, Fiber, Stream } from "effect"
import * as Scheduler from "effect/Scheduler"
import { afterEach, describe, expect, it, vi } from "vitest"
import * as Workbook from "../src/Workbook.ts"
import { isDefaultScheduler, messageChannelTask, type PostTask, yieldFor, yieldThrough } from "../src/internal/yield.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { bigXlsx } from "./helpers/big-xlsx.ts"

/** A post that runs its tasks on setImmediate and counts them. */
const countingPost = () => {
  const counter = { posts: 0 }
  const post: PostTask = (task) => {
    counter.posts++
    setImmediate(task)
  }
  return { counter, post }
}

const yieldTimes = (yieldOnce: Effect.Effect<void>, n: number) =>
  Effect.gen(function*() {
    // progress: i up to n
    for (let i = 0; i < n; i++) yield* yieldOnce
    return n
  })

describe("reader yields", () => {
  afterEach(() => {
    vi.unstubAllGlobals()
    vi.resetModules()
  })

  it("know Effect's default scheduler from one the application chose", () => {
    expect(isDefaultScheduler(new Scheduler.MixedScheduler())).toBe(true)
    expect(isDefaultScheduler(new Scheduler.MixedScheduler("async"))).toBe(true)
    expect(isDefaultScheduler(new Scheduler.MixedScheduler("sync"))).toBe(false)
    expect(isDefaultScheduler(new Scheduler.MixedScheduler("async", (f) => {
      const timer = setImmediate(f)
      return () => clearImmediate(timer)
    }))).toBe(false)
    const own: Scheduler.Scheduler = {
      executionMode: "async",
      shouldYield: () => false,
      makeDispatcher: () => new Scheduler.MixedScheduler().makeDispatcher()
    }
    expect(isDefaultScheduler(own)).toBe(false)
  })

  it("are Effect.yieldNow where there is setImmediate or no MessageChannel, else their own", () => {
    const noop = () => undefined
    expect(yieldFor({ setImmediate: noop, MessageChannel: noop })).toBe(Effect.yieldNow)
    expect(yieldFor({ setImmediate: noop })).toBe(Effect.yieldNow)
    expect(yieldFor({})).toBe(Effect.yieldNow)
    expect(yieldFor({ setImmediate: 1, MessageChannel: 1 })).toBe(Effect.yieldNow)
    expect(yieldFor({ MessageChannel: noop })).not.toBe(Effect.yieldNow)
  })

  it("resume through the posted task on the default scheduler, once per yield", async () => {
    const { counter, post } = countingPost()
    expect(await Effect.runPromise(yieldTimes(yieldThrough(post), 50))).toBe(50)
    expect(counter.posts).toBe(50)
  })

  it("leave a scheduler the application provided in charge", async () => {
    const { counter, post } = countingPost()
    let dispatched = 0
    const scheduler = new Scheduler.MixedScheduler("async", (f) => {
      dispatched++
      const timer = setImmediate(f)
      return () => clearImmediate(timer)
    })
    const run = yieldTimes(yieldThrough(post), 50).pipe(Effect.provideService(Scheduler.Scheduler, scheduler))
    expect(await Effect.runPromise(run)).toBe(50)
    expect(counter.posts).toBe(0)
    expect(dispatched).toBeGreaterThanOrEqual(50)
  })

  it("can be interrupted while waiting; the late task then does nothing", async () => {
    const held: Array<() => void> = []
    let after = 0
    const program = Effect.gen(function*() {
      yield* yieldThrough((task) => void held.push(task))
      after++
    })
    const exit = await Effect.runPromise(Effect.gen(function*() {
      const fiber = yield* Effect.forkChild(program)
      yield* Effect.yieldNow
      expect(held.length).toBe(1)
      return yield* Fiber.interrupt(fiber).pipe(Effect.andThen(Fiber.await(fiber)))
    }))
    expect(Exit.hasInterrupts(exit)).toBe(true)
    held[0]!()
    await new Promise((resolve) => setImmediate(resolve))
    expect(after).toBe(0)
  })

  it("post MessageChannel tasks in order, each run once, all through one channel", async () => {
    let channels = 0
    class CountedChannel extends MessageChannel {
      constructor() {
        super()
        channels++
      }
    }
    vi.stubGlobal("MessageChannel", CountedChannel)
    const post = messageChannelTask()
    const ran: Array<number> = []
    await new Promise<void>((resolve) => {
      post(() => ran.push(1))
      post(() => ran.push(2))
      post(() => {
        ran.push(3)
        post(() => {
          ran.push(4)
          resolve()
        })
      })
    })
    expect(ran).toEqual([1, 2, 3, 4])
    expect(channels).toBe(1)
  })

  it("take MessageChannel messages in a runtime without setImmediate, and read the same rows", async () => {
    const bytes = await bigXlsx(20_000)
    const readRows = (xlsx: typeof Xlsx, workbook: typeof Workbook) =>
      Effect.runPromise(Effect.gen(function*() {
        const wb = yield* xlsx.read(bytes)
        const out = yield* Stream.runCollect(workbook.rows(wb, 0))
        return Array.from(out)
      }))
    const expected = await readRows(Xlsx, Workbook)

    // a stand-in channel that delivers on setImmediate and counts the messages
    let messages = 0
    const deliver = setImmediate
    class CountingChannel {
      readonly port1: { onmessage: (() => void) | null } = { onmessage: null }
      readonly port2 = {
        postMessage: () => {
          messages++
          deliver(() => this.port1.onmessage?.())
        }
      }
    }
    vi.stubGlobal("setImmediate", undefined)
    vi.stubGlobal("MessageChannel", CountingChannel)
    vi.resetModules()
    const freshXlsx = await import("../src/xlsx/index.ts")
    const freshWorkbook = await import("../src/Workbook.ts")
    // the channel is made on the first yield, so the stand-in stays until the read is done
    const rows = await readRows(freshXlsx, freshWorkbook)
    expect(rows).toEqual(expected)
    // one yield per 64 KiB slice of the sheet XML and more
    expect(messages).toBeGreaterThan(10)
  })
})

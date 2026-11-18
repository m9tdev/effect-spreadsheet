import { describe, expect, it } from "vitest"
import { inflateRawWeb } from "../src/internal/zip/inflate.ts"
import { deflateRaw } from "./helpers/zip.ts"

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

/** How a test brings the inflater to the moment it cancels. */
interface Setup {
  /** Waits until the pump has filled its hold and the stream's source has paused. */
  readonly settle: () => Promise<void>
  /** How often the stream has pulled from its source, when the test can see it. */
  readonly pulls?: () => number
}

/**
 * Lets an inflater fill its hold, takes a chunk that makes the pump read
 * again (so the stream pulls from its source), and cancels at once, before
 * any task runs. An exception the stream throws afterwards is uncaught: it
 * is collected here (the test run reports it as well).
 */
const cancelRightAfterARead = async (compressed: Uint8Array, setup: Setup) => {
  const uncaught: Array<unknown> = []
  const collect = (error: unknown) => {
    uncaught.push(error)
  }
  process.on("uncaughtException", collect)
  try {
    await readAndCancel(compressed, setup)
  } finally {
    process.off("uncaughtException", collect)
  }
  expect(uncaught).toEqual([])
}

const readAndCancel = async (compressed: Uint8Array, { pulls, settle }: Setup) => {
  const inflater = inflateRawWeb(compressed)
  expect(await inflater.next()).toBeDefined()
  await settle()
  if (pulls === undefined) {
    expect(await inflater.next()).toBeDefined()
  } else {
    // Take held chunks until one brings the hold under its cap: the pump
    // then reads, and the stream pulls. No task runs in between, since each
    // next() here returns a held chunk.
    const before = pulls()
    // progress: taken is bounded
    for (let taken = 0; pulls() === before && taken < 1000; taken++) {
      expect(await inflater.next()).toBeDefined()
    }
    expect(pulls()).toBeGreaterThan(before)
  }
  inflater.cancel()
  for (let i = 0; i < 20; i++) await tick()
  // the output taken before the cancel is still handed out, then the end
  let reads = 0
  // progress: reads is bounded; each next() returns a held chunk or the end
  while (reads < 10_000 && (await inflater.next()) !== undefined) reads++
  expect(reads).toBeLessThan(10_000)
}

describe("cancelling an inflater while the stream pulls from its source", () => {
  it("throws nothing with the platform's DecompressionStream", async () => {
    // 8 MiB of zeros: one slice that inflates past the hold, so the pump pauses
    await cancelRightAfterARead(await deflateRaw(new Uint8Array(8 * 1024 * 1024)), {
      settle: async () => {
        for (let i = 0; i < 50; i++) await tick()
      }
    })
  })

  it("throws nothing when the source hands buffered output to the stream in a callback scheduled at pull, unchecked", async () => {
    // Models Node's adapter before 24.22: the decompressor's output waits in
    // its own buffer while the stream is paused; a pull schedules a callback
    // that enqueues that buffer without checking whether the stream was
    // cancelled since. Output produced after the cancel is dropped.
    const Real = globalThis.DecompressionStream
    let last: UncheckedResume | undefined
    class UncheckedResume {
      readonly writable: WritableStream<BufferSource>
      readonly readable: ReadableStream<Uint8Array>
      pulls = 0
      /** Paused with output waiting: it filled the stream's queue after a pull and holds more. */
      paused = false
      constructor(format: CompressionFormat) {
        last = this
        const inner = new Real(format)
        const source = inner.readable.getReader()
        this.writable = inner.writable
        const buffered: Array<Uint8Array> = []
        let flowing = false
        let ended = false
        let cancelled = false
        let controller!: ReadableStreamDefaultController<Uint8Array>
        const flush = () => {
          // progress: each iteration hands out one buffered chunk
          while (flowing && buffered.length > 0) {
            controller.enqueue(buffered.shift()!)
            if ((controller.desiredSize ?? 0) <= 0) flowing = false
          }
          this.paused = this.pulls > 0 && !flowing && buffered.length > 0
          if (ended && buffered.length === 0) {
            ended = false
            controller.close()
          }
        }
        void (async () => {
          // progress: each read takes a chunk or ends the loop
          while (!cancelled) {
            const result = await source.read().catch(() => ({ done: true as const, value: undefined }))
            if (cancelled) return
            if (result.done) ended = true
            else buffered.push(result.value)
            flush()
            if (result.done) return
          }
        })()
        this.readable = new ReadableStream<Uint8Array>({
          start: (c) => {
            controller = c
          },
          pull: () => {
            this.pulls++
            process.nextTick(() => {
              flowing = true
              flush()
            })
          },
          cancel: (reason) => {
            cancelled = true
            return source.cancel(reason)
          }
        }, new ByteLengthQueuingStrategy({ highWaterMark: 16 * 1024 }))
      }
    }
    globalThis.DecompressionStream = UncheckedResume as unknown as typeof DecompressionStream
    try {
      await cancelRightAfterARead(await deflateRaw(new Uint8Array(8 * 1024 * 1024)), {
        // Between tasks, a paused source means the stream's queue is full, so
        // the pump has stopped reading at its cap.
        settle: async () => {
          // progress: i is bounded
          for (let i = 0; i < 5000 && last?.paused !== true; i++) await tick()
          expect(last?.paused).toBe(true)
        },
        pulls: () => last?.pulls ?? 0
      })
    } finally {
      globalThis.DecompressionStream = Real
    }
  })
})

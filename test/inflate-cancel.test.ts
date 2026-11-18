import { describe, expect, it } from "vitest"
import { inflateRawWeb } from "../src/internal/zip/inflate.ts"
import { deflateRaw } from "./helpers/zip.ts"

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

/**
 * Lets an inflater fill its hold, takes one more chunk (the pump then reads
 * again, which makes the stream pull from its source), and cancels at once,
 * before any task runs. An exception the stream throws afterwards is
 * uncaught: it is collected here (the test run reports it as well).
 */
const cancelRightAfterARead = async (compressed: Uint8Array) => {
  const uncaught: Array<unknown> = []
  const collect = (error: unknown) => {
    uncaught.push(error)
  }
  process.on("uncaughtException", collect)
  try {
    await readAndCancel(compressed)
  } finally {
    process.off("uncaughtException", collect)
  }
  expect(uncaught).toEqual([])
}

const readAndCancel = async (compressed: Uint8Array) => {
  const inflater = inflateRawWeb(compressed)
  expect(await inflater.next()).toBeDefined()
  // let the pump fill its hold and the stream's queue, so its source pauses
  for (let i = 0; i < 50; i++) await tick()
  expect(await inflater.next()).toBeDefined()
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
    await cancelRightAfterARead(await deflateRaw(new Uint8Array(8 * 1024 * 1024)))
  })

  it("throws nothing when the source hands buffered output to the stream in a callback scheduled at pull, unchecked", async () => {
    // Models Node's adapter before 24.22: the decompressor's output waits in
    // its own buffer while the stream is paused; a pull schedules a callback
    // that enqueues that buffer without checking whether the stream was
    // cancelled since. Output produced after the cancel is dropped.
    const Real = globalThis.DecompressionStream
    class UncheckedResume {
      readonly writable: WritableStream<BufferSource>
      readonly readable: ReadableStream<Uint8Array>
      constructor(format: CompressionFormat) {
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
      await cancelRightAfterARead(await deflateRaw(new Uint8Array(8 * 1024 * 1024)))
    } finally {
      globalThis.DecompressionStream = Real
    }
  })
})

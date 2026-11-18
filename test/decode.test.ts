import { Effect, Exit, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { decoded, drained, type PushDecoder, step, stopped } from "../src/internal/decode.ts"
import type { Meter } from "../src/internal/meter.ts"
import { MalformedFile } from "../src/SpreadsheetError.ts"
import { counter, metered } from "./helpers/meter.ts"
import { failureOf } from "./helpers/run.ts"

const bad = (offset: number) => new MalformedFile({ format: "xml", offset, reason: "bad byte" })

/**
 * A toy decoder: each byte below 0xF0 completes one item (its value); 0xFF
 * is an error after the items before it in the piece; the end emits -1,
 * then fails if `failAtEnd`. Counts bytes on the meter it is given.
 */
const toy = (log: Array<string>, failAtEnd = false) => (meter: Meter): PushDecoder<number> => {
  let offset = 0
  log.push("make")
  return {
    feed: (piece) => {
      log.push(`feed ${Array.from(piece).join(",")}`)
      meter.add("bytes", piece.length)
      const out: Array<number> = []
      for (const byte of piece) {
        if (byte === 0xFF) return step(out, bad(offset))
        if (byte < 0xF0) out.push(byte)
        offset++
      }
      return step(out, undefined)
    },
    end: () => {
      log.push("end")
      return step([-1], failAtEnd ? bad(offset) : undefined)
    }
  }
}

/** The same decoder, keeping only its errors (for `drained`). */
const quiet = (make: (meter: Meter) => PushDecoder<number>) => (meter: Meter): PushDecoder<never> => {
  const decoder = make(meter)
  return {
    feed: (piece) => stopped(decoder.feed(piece).error),
    end: () => stopped(decoder.end().error)
  }
}

const bytes = (...values: Array<number>) => Uint8Array.from(values)

describe("decoded", () => {
  it("emits what each upstream chunk completes, one emission per chunk, and none for a chunk that completes nothing", async () => {
    const log: Array<string> = []
    const input = Stream.fromArrays([bytes(1, 2), bytes(3)], [bytes(0xF0)], [bytes(4)])
    const chunks = await Effect.runPromise(Stream.runCollect(Stream.chunks(decoded(input, toy(log)))))
    expect(Array.from(chunks, (chunk) => Array.from(chunk))).toEqual([[1, 2, 3], [4], [-1]])
    expect(log).toEqual(["make", "feed 1,2", "feed 3", "feed 240", "feed 4", "end"])
  })

  it("emits what came before an error, then fails with it, and reads no further", async () => {
    const log: Array<string> = []
    const seen: Array<number> = []
    const input = Stream.fromArrays([bytes(1)], [bytes(2, 0xFF, 3), bytes(5)], [bytes(6)])
    const exit = await Effect.runPromiseExit(
      Stream.runForEach(decoded(input, toy(log)), (n) => Effect.sync(() => void seen.push(n)))
    )
    expect(seen).toEqual([1, 2])
    expect(failureOf(exit)).toEqual(bad(2))
    expect(log).toEqual(["make", "feed 1", "feed 2,255,3"])
  })

  it("emits what the end completes before the end's error", async () => {
    const seen: Array<number> = []
    const exit = await Effect.runPromiseExit(
      Stream.runForEach(decoded(Stream.make(bytes(7)), toy([], true)), (n) => Effect.sync(() => void seen.push(n)))
    )
    expect(seen).toEqual([7, -1])
    expect(failureOf(exit)).toEqual(bad(1))
  })

  it("passes an input error through at once", async () => {
    const log: Array<string> = []
    const input = Stream.concat(Stream.make(bytes(1)), Stream.fail("input"))
    const exit = await Effect.runPromiseExit(Stream.runCollect(decoded(input, toy(log))))
    expect(failureOf(exit)).toBe("input")
    expect(log).toEqual(["make", "feed 1"])
  })

  it("makes a fresh decoder per run, with the meter of that run", async () => {
    const log: Array<string> = []
    const stream = decoded(Stream.make(bytes(1, 2), bytes(3)), toy(log))
    const meter = counter("bytes")
    const first = await Effect.runPromise(metered(meter, Stream.runCollect(stream)))
    const second = await Effect.runPromise(Stream.runCollect(stream))
    expect(Array.from(first)).toEqual([1, 2, 3, -1])
    expect(Array.from(second)).toEqual(Array.from(first))
    expect(meter.total()).toBe(3)
    expect(log.filter((line) => line === "make")).toHaveLength(2)
  })
})

describe("drained", () => {
  it("feeds every piece, then ends", async () => {
    const log: Array<string> = []
    const exit = await Effect.runPromiseExit(drained(Stream.fromArrays([bytes(1), bytes(2)]), quiet(toy(log))))
    expect(Exit.isSuccess(exit)).toBe(true)
    expect(log).toEqual(["make", "feed 1", "feed 2", "end"])
  })

  it("fails with the first error and reads no further", async () => {
    const log: Array<string> = []
    const exit = await Effect.runPromiseExit(drained(Stream.make(bytes(1, 0xFF), bytes(2)), quiet(toy(log))))
    expect(failureOf(exit)).toEqual(bad(1))
    expect(log).toEqual(["make", "feed 1,255"])
  })

  it("fails with the end's error", async () => {
    const exit = await Effect.runPromiseExit(drained(Stream.make(bytes(1)), quiet(toy([], true))))
    expect(failureOf(exit)).toEqual(bad(1))
  })
})

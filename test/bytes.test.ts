import { Effect, Stream } from "effect"
import { describe, expect, it } from "vitest"
import { SLICE_BYTES, sliced } from "../src/internal/bytes.ts"
import { countYields } from "./helpers/yields.ts"
import { concat } from "./helpers/zip.ts"

const bytes = (n: number, seed: number): Uint8Array => Uint8Array.from({ length: n }, (_, i) => (i * 31 + seed) & 0xFF)

describe("sliced byte streams", () => {
  it("hand out one slice of at most SLICE_BYTES per pull, each after a yield; empty pieces are dropped", async () => {
    const pieces = [bytes(3 * SLICE_BYTES + 5, 1), new Uint8Array(0), bytes(10, 2), bytes(SLICE_BYTES, 3)]
    // two pieces in one upstream chunk, then two more chunks
    const input = Stream.fromArrays([pieces[0]!, pieces[1]!], [pieces[2]!], [pieces[3]!])
    const run = await countYields(() => Effect.runPromise(Stream.runCollect(Stream.chunks(sliced(input)))))
    const pulls = Array.from(run.value)
    expect(pulls.every((pull) => pull.length === 1)).toBe(true)
    const slices = pulls.map((pull) => pull[0]!)
    expect(slices.map((slice) => slice.length)).toEqual([SLICE_BYTES, SLICE_BYTES, SLICE_BYTES, 5, 10, SLICE_BYTES])
    expect(Buffer.from(concat(slices)).equals(concat(pieces))).toBe(true)
    expect(run.yields).toBeGreaterThanOrEqual(slices.length / 2)
  })

  it("pull an upstream chunk only once the slices of the one before it are pulled", async () => {
    let pulled = 0
    const input = Stream.fromArrays([bytes(2 * SLICE_BYTES, 1)], [bytes(10, 2)]).pipe(
      Stream.tap(() => Effect.sync(() => void pulled++))
    )
    const seen: Array<number> = []
    await Effect.runPromise(Stream.runForEach(sliced(input), () => Effect.sync(() => void seen.push(pulled))))
    expect(seen).toEqual([1, 1, 2])
  })
})

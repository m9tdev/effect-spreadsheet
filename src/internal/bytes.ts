/**
 * Byte streams in bounded slices, with one yield to the event loop before
 * each slice. Readers parse what they are given as it comes and do no
 * slicing or yielding of their own: the bound on the work between two
 * turns of the loop comes from the byte source. The sources are the zip
 * archive's entry reads (which cut what they inflate the same way) and
 * `sliced`, for byte streams from anywhere else (a whole buffer, a user's
 * stream).
 *
 * @internal
 */
import { Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"

/** Most bytes handed to a reader between two event-loop turns. */
export const SLICE_BYTES = 64 * 1024

/**
 * The stream as one slice of at most `SLICE_BYTES` per pull, each pulled
 * after a yield. Empty pieces are dropped. An upstream chunk is only
 * pulled once the slices of the one before it have all been pulled.
 */
export const sliced = <E, R>(bytes: Stream.Stream<Uint8Array, E, R>): Stream.Stream<Uint8Array, E, R> =>
  Stream.transformPull(bytes, (upstream) =>
    Effect.sync(() => {
      /** The upstream chunk being cut, the piece in it and the offset in that piece. */
      let pieces: ReadonlyArray<Uint8Array> = []
      let index = 0
      let at = 0
      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Uint8Array>, E> = Effect.gen(function*() {
        // progress: each iteration hands out a non-empty slice (at grows),
        // moves past a used-up piece (index grows), or pulls one upstream
        // chunk (consumed); upstream's end ends the stream
        while (true) {
          const piece = pieces[index]
          if (piece === undefined) {
            pieces = yield* upstream
            index = 0
            at = 0
            continue
          }
          if (at >= piece.length) {
            index++
            at = 0
            continue
          }
          yield* Effect.yieldNow
          const slice = piece.subarray(at, Math.min(piece.length, at + SLICE_BYTES))
          at += slice.length
          return [slice] as const
        }
      })
      return pull
    }))

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
import { Effect, Stream } from "effect"

/** Most bytes handed to a reader between two event-loop turns. */
export const SLICE_BYTES = 64 * 1024

/** `piece` cut into non-empty slices of at most `SLICE_BYTES` (views, not copies). */
export const slicesOf = (piece: Uint8Array): Array<Uint8Array> => {
  if (piece.length <= SLICE_BYTES) return piece.length === 0 ? [] : [piece]
  const out: Array<Uint8Array> = []
  // progress: at grows by SLICE_BYTES up to piece.length
  for (let at = 0; at < piece.length; at += SLICE_BYTES) out.push(piece.subarray(at, Math.min(piece.length, at + SLICE_BYTES)))
  return out
}

/**
 * The stream as one slice of at most `SLICE_BYTES` per pull, each pulled
 * after a yield. Empty pieces are dropped. An upstream chunk is only
 * pulled once the slices of the one before it have all been pulled.
 */
export const sliced = <E, R>(bytes: Stream.Stream<Uint8Array, E, R>): Stream.Stream<Uint8Array, E, R> =>
  bytes.pipe(
    Stream.flatMap((piece) => Stream.fromIterable(slicesOf(piece))),
    Stream.rechunk(1),
    Stream.tap(() => Effect.yieldNow)
  )

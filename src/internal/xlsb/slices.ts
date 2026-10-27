/**
 * A part's byte stream in bounded slices, with a yield to the event loop
 * before each one: whatever size the pieces come in (the inflater's
 * output, a stored entry's slices), the records split between two turns
 * of the loop come from at most `SLICE_BYTES` bytes.
 *
 * @internal
 */
import { Effect, Stream } from "effect"

/** Bytes handled between event-loop turns. */
export const SLICE_BYTES = 64 * 1024

/** `piece` cut into slices of at most `SLICE_BYTES` (views, not copies). */
export const slicesOf = (piece: Uint8Array): Array<Uint8Array> => {
  if (piece.length <= SLICE_BYTES) return [piece]
  const out: Array<Uint8Array> = []
  // progress: at grows by SLICE_BYTES up to piece.length
  for (let at = 0; at < piece.length; at += SLICE_BYTES) out.push(piece.subarray(at, Math.min(piece.length, at + SLICE_BYTES)))
  return out
}

/** The stream in slices of at most `SLICE_BYTES`, each pulled after a yield. */
export const sliced = <E, R>(bytes: Stream.Stream<Uint8Array, E, R>): Stream.Stream<Uint8Array, E, R> =>
  bytes.pipe(
    Stream.flatMap((piece) => Stream.fromIterable(slicesOf(piece))),
    Stream.rechunk(1),
    Stream.tap(() => Effect.yieldNow)
  )

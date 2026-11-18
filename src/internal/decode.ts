/**
 * Push decoders, and the one driver that runs them over a byte stream.
 *
 * A part's bytes are decoded as they come: a push decoder is fed each
 * piece, then told the input ended, and each call says what the bytes
 * read so far completed (rows of a sheet, say; nothing for a decoder that
 * works through a handler) and whether they hit an error. `decoded` turns
 * one into a stream of what it completes, `drained` runs one for its
 * effects. The byte source hands over bounded slices, each after a yield
 * to the event loop (`bytes.ts`), so the work between two turns of the
 * loop is bounded without the driver slicing anything itself.
 *
 * @internal
 */
import { Cause, Effect, Pull, Stream } from "effect"
import * as Arr from "effect/Array"
import { absurd } from "effect/Function"
import type { SpreadsheetError } from "../SpreadsheetError.ts"
import { type Meter, WorkMeter } from "./meter.ts"

/** What a step of a decoder completed, and the error that stopped it, if any. */
export interface Step<A> {
  readonly out: ReadonlyArray<A>
  readonly error?: SpreadsheetError
}

/**
 * A decoder fed a part's bytes: `feed` for each piece, in order, then `end`
 * once the input ends. After a step with an error it isn't called again.
 */
export interface PushDecoder<A> {
  feed(piece: Uint8Array): Step<A>
  end(): Step<A>
}

/** A step that completed nothing and found no error. */
const NOTHING: Step<never> = { out: [] }

/** A step with `out`, failed with `error` if there is one. */
export const step = <A>(out: ReadonlyArray<A>, error: SpreadsheetError | undefined): Step<A> =>
  error === undefined ? { out } : { out, error }

/** A step that completed nothing, failed with `error` if there is one. */
export const stopped = (error: SpreadsheetError | undefined): Step<never> =>
  error === undefined ? NOTHING : { out: [], error }

/**
 * Where a decoded stream is: reading input; failing, with what the input
 * completed before the error already emitted; or done (the input ended, or
 * the error was raised).
 */
type State =
  | { readonly _tag: "Reading" }
  | { readonly _tag: "Failing"; readonly error: SpreadsheetError }
  | { readonly _tag: "Done" }

const READING: State = { _tag: "Reading" }
const DONE: State = { _tag: "Done" }

/** Feeds the pieces of one chunk, up to the first error; one step for them all. */
const feedAll = <A>(decoder: PushDecoder<A>, pieces: ReadonlyArray<Uint8Array>): Step<A> => {
  const out: Array<A> = []
  // progress: one piece per iteration
  for (const piece of pieces) {
    const next = decoder.feed(piece)
    // progress: one item per iteration
    for (const a of next.out) out.push(a)
    if (next.error !== undefined) return { out, error: next.error }
  }
  return { out }
}

/**
 * What the decoder made by `make` completes from `bytes`, emitted with the
 * upstream chunk that completed it. On an error (from the decoder), what
 * the input completed before it comes out first, then the stream fails
 * with it; so what comes out before an error doesn't depend on where the
 * input is split. An error from `bytes` fails the stream at once. `make`
 * runs for each run of the stream, with the work meter in effect there, so
 * the stream is re-runnable when the input is.
 */
export const decoded = <A, E, R>(
  bytes: Stream.Stream<Uint8Array, E, R>,
  make: (meter: Meter) => PushDecoder<A>
): Stream.Stream<A, E | SpreadsheetError, R> =>
  Stream.transformPull(bytes, (upstream) =>
    Effect.map(WorkMeter, (meter) => {
      const decoder = make(meter)
      const pullUpstream = Pull.catchDone(upstream, () => Effect.undefined)
      let state: State = READING
      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<A>, E | SpreadsheetError> = Effect.gen(function*() {
        // progress: each iteration pulls one upstream chunk (consumed), or
        // ends the input (Done), or ends the stream (Failing goes to Done)
        while (true) {
          switch (state._tag) {
            case "Reading":
              break
            case "Failing": {
              const error = state.error
              state = DONE
              return yield* Effect.fail(error)
            }
            case "Done":
              return yield* Cause.done()
            default:
              return absurd(state)
          }
          const chunk = yield* pullUpstream
          let next: Step<A>
          if (chunk === undefined) {
            state = DONE
            next = decoder.end()
          } else {
            next = chunk.length === 1 ? decoder.feed(chunk[0]) : feedAll(decoder, chunk)
          }
          // what came before the error comes out first; the error on the next pull
          if (next.error !== undefined) state = { _tag: "Failing", error: next.error }
          const out = next.out
          if (Arr.isReadonlyArrayNonEmpty(out)) return out
        }
      })
      return pull
    }))

/**
 * Runs the decoder made by `make` over `bytes`, for its effects. Fails
 * with the first error, from `bytes` or the decoder. `make` runs for each
 * run, with the work meter in effect there.
 */
export const drained = <E, R>(
  bytes: Stream.Stream<Uint8Array, E, R>,
  make: (meter: Meter) => PushDecoder<never>
): Effect.Effect<void, E | SpreadsheetError, R> =>
  Effect.flatMap(WorkMeter, (meter) => {
    const decoder = make(meter)
    return Stream.runForEach(bytes, (piece) => {
      const error = decoder.feed(piece).error
      return error === undefined ? Effect.void : Effect.fail(error)
    }).pipe(
      Effect.flatMap(() => {
        const error = decoder.end().error
        return error === undefined ? Effect.void : Effect.fail(error)
      })
    )
  })

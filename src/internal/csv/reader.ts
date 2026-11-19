/**
 * Byte stream -> row stream for delimited text.
 *
 * @internal
 */
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import * as Result from "effect/Result"
import { type InvalidOption, LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
import { sliced } from "../bytes.ts"
import { csvText } from "./decode.ts"
import type { CsvOptions } from "./options.ts"
import { CsvParser } from "./parser.ts"

/**
 * The input, failing with `LimitExceeded` once its bytes add up past
 * `maxInputBytes`: counted per upstream chunk, before any of that chunk is
 * handed on.
 */
const limitInput = <E, R>(input: Stream.Stream<Uint8Array, E, R>, maxInputBytes: number): Stream.Stream<Uint8Array, E | LimitExceeded, R> =>
  Stream.transformPull(input, (upstream) =>
    Effect.sync(() => {
      let inputBytes = 0
      return Effect.flatMap(upstream, (chunk) => {
        // progress: one piece per iteration
        for (const c of chunk) inputBytes += c.length
        return inputBytes > maxInputBytes
          ? Effect.fail(new LimitExceeded({ limit: "maxInputBytes", max: maxInputBytes, actual: inputBytes }))
          : Effect.succeed(chunk)
      })
    }))

/**
 * Parses a byte stream into rows, with checked options and the limits in
 * effect. The input is read in slices of at most 64 KiB, each decoded and
 * parsed after a yield to the event loop.
 */
export const rowsFromByteStream = <E, R>(
  input: Stream.Stream<Uint8Array, E, R>,
  options: CsvOptions,
  limits: Limits
): Stream.Stream<Row, E | SpreadsheetError, R> =>
  Stream.transformPull(sliced(limitInput(input, limits.maxInputBytes)), (upstream) =>
    Effect.sync(() => {
      const decoder = csvText(options.decoding)
      const parser = new CsvParser({
        delimiter: options.delimiter,
        inferTypes: options.inferTypes,
        strict: options.strict,
        limits
      })
      let finished = false
      const pullUpstream = Pull.catchDone(upstream, () => Effect.undefined)

      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, E | SpreadsheetError> = Effect.gen(function*() {
        // progress: each iteration pulls one slice (consumed) or finishes.
        // After `finished` the pull ends.
        while (true) {
          // rows completed before a parse error came out on the last pull;
          // the error comes now, before any more input is read or decoded
          if (parser.error !== undefined) return yield* Effect.fail(parser.error)
          if (finished) return yield* Cause.done()
          let rows: Result.Result<Array<Row>, SpreadsheetError>
          const next = yield* pullUpstream
          if (next !== undefined) {
            let text = ""
            // progress: one slice per iteration (one per pull, as `sliced` emits them)
            for (const slice of next) text += Result.merge(decoder.push(slice))
            rows = parser.feed(text, false)
          } else {
            finished = true
            rows = parser.feed(Result.merge(decoder.end()), true)
          }
          if (Result.isFailure(rows)) return yield* Effect.fail(rows.failure)
          if (rows.success.length > 0) return rows.success as unknown as Arr.NonEmptyReadonlyArray<Row>
        }
      })
      return pull
    }))

/**
 * What `Csv.rows` streams when the requested encoding is unknown: the input
 * is read as far as decoding would read it before choosing the decoder (3
 * bytes, to look for a BOM, or the whole input if shorter), then the stream
 * fails with `error`. So the input's own errors and the size limit, met
 * before that point, come first, as they do with a known encoding.
 */
export const failAtDecoding = <E, R>(
  input: Stream.Stream<Uint8Array, E, R>,
  limits: Limits,
  error: InvalidOption
): Stream.Stream<Row, E | SpreadsheetError, R> =>
  Stream.transformPull(sliced(limitInput(input, limits.maxInputBytes)), (upstream) =>
    Effect.sync(() => {
      const pullUpstream = Pull.catchDone(upstream, () => Effect.undefined)
      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, E | SpreadsheetError> = Effect.gen(function*() {
        let bytes = 0
        // progress: each iteration pulls a non-empty slice (`sliced` drops
        // empty pieces), so `bytes` grows, or meets the input's end and stops
        while (bytes < 3) {
          const next = yield* pullUpstream
          if (next === undefined) break
          for (const slice of next) bytes += slice.length
        }
        return yield* Effect.fail(error)
      })
      return pull
    }))

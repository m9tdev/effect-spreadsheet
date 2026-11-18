/**
 * Byte stream -> row stream for delimited text.
 *
 * @internal
 */
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import * as Result from "effect/Result"
import { InvalidOption, LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
import { sliced } from "../bytes.ts"
import { ByteDecoder } from "./decode.ts"
import { CsvParser } from "./parser.ts"

export interface ResolvedOptions {
  readonly encoding: string
  readonly delimiter: string | undefined
  readonly inferTypes: boolean
  readonly strict: boolean
  readonly limits: Limits
}

/**
 * Checks `options.delimiter` at runtime, since the type can be bypassed
 * (plain JavaScript, casts): `undefined` (guess it) or a one-character
 * string other than a quote, CR or LF.
 */
export const validateDelimiter = (delimiter: unknown): InvalidOption | undefined => {
  if (delimiter === undefined) return undefined
  if (typeof delimiter !== "string") {
    const type = typeof delimiter
    const kind = delimiter === null
      ? "null"
      : Array.isArray(delimiter)
      ? "an array"
      : type === "object"
      ? "an object"
      : `a ${type}`
    return new InvalidOption({ option: "delimiter", reason: `must be a one-character string, not ${kind}` })
  }
  return delimiter.length === 1 && delimiter !== '"' && delimiter !== "\r" && delimiter !== "\n"
    ? undefined
    : new InvalidOption({
      option: "delimiter",
      reason: `${JSON.stringify(delimiter)} must be one character other than a quote, CR or LF`
    })
}

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
 * Parses a byte stream into rows. The input is read in slices of at most
 * 64 KiB, each decoded and parsed after a yield to the event loop.
 */
export const rowsFromByteStream = <E, R>(
  input: Stream.Stream<Uint8Array, E, R>,
  options: ResolvedOptions
): Stream.Stream<Row, E | SpreadsheetError, R> => {
  const invalid = validateDelimiter(options.delimiter)
  if (invalid !== undefined) return Stream.fail(invalid)
  return Stream.transformPull(sliced(limitInput(input, options.limits.maxInputBytes)), (upstream) =>
    Effect.sync(() => {
      const decoder = new ByteDecoder(options.encoding)
      const parser = new CsvParser({
        delimiter: options.delimiter,
        inferTypes: options.inferTypes,
        strict: options.strict,
        limits: options.limits
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
            for (const slice of next) {
              const decoded = decoder.push(slice)
              if (Result.isFailure(decoded)) return yield* Effect.fail(decoded.failure)
              text += decoded.success
            }
            rows = parser.feed(text, false)
          } else {
            finished = true
            const text = decoder.end()
            if (Result.isFailure(text)) return yield* Effect.fail(text.failure)
            rows = parser.feed(text.success, true)
          }
          if (Result.isFailure(rows)) return yield* Effect.fail(rows.failure)
          if (rows.success.length > 0) return rows.success as unknown as Arr.NonEmptyReadonlyArray<Row>
        }
      })
      return pull
    }))
}

/**
 * Byte stream -> row stream for delimited text.
 *
 * @internal
 */
import { Cause, Effect, Pull, Stream } from "effect"
import type * as Arr from "effect/Array"
import { InvalidOption, LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import type { Row } from "../../Workbook.ts"
import { ByteDecoder } from "./decode.ts"
import { CsvParser } from "./parser.ts"

/**
 * Bytes decoded and parsed per step. Each step is bounded work; the fiber
 * yields to the event loop before each one.
 */
export const PIECE_BYTES = 64 * 1024

export interface ResolvedOptions {
  readonly encoding: string
  readonly delimiter: string | undefined
  readonly inferTypes: boolean
  readonly strict: boolean
  readonly limits: Limits
}

export const validateDelimiter = (delimiter: string | undefined): InvalidOption | undefined =>
  delimiter === undefined || (delimiter.length === 1 && delimiter !== '"' && delimiter !== "\r" && delimiter !== "\n")
    ? undefined
    : new InvalidOption({
      option: "delimiter",
      reason: `${JSON.stringify(delimiter)} must be one character other than a quote, CR or LF`
    })

const isError = (u: Array<Row> | SpreadsheetError): u is SpreadsheetError => !Array.isArray(u)

/** Parses a byte stream into rows. */
export const rowsFromByteStream = <E, R>(
  input: Stream.Stream<Uint8Array, E, R>,
  options: ResolvedOptions
): Stream.Stream<Row, E | SpreadsheetError, R> => {
  const invalid = validateDelimiter(options.delimiter)
  if (invalid !== undefined) return Stream.fail(invalid)
  return Stream.transformPull(input, (upstream) =>
    Effect.sync(() => {
      const decoder = new ByteDecoder(options.encoding)
      const parser = new CsvParser({
        delimiter: options.delimiter,
        inferTypes: options.inferTypes,
        strict: options.strict,
        limits: options.limits
      })
      const maxInputBytes = options.limits.maxInputBytes
      let chunks: ReadonlyArray<Uint8Array> = []
      let chunkIndex = 0
      let offset = 0
      let inputBytes = 0
      let upstreamDone = false
      let finished = false
      const pullUpstream = Pull.catchDone(upstream, () => Effect.succeed(undefined))

      const pull: Pull.Pull<Arr.NonEmptyReadonlyArray<Row>, E | SpreadsheetError> = Effect.gen(function*() {
        // progress: each iteration consumes >= 1 queued byte (offset grows,
        // chunkIndex grows past exhausted chunks), or pulls one upstream
        // chunk, or finishes. After `finished` the pull ends.
        while (true) {
          if (finished) return yield* Cause.done()
          let rows: Array<Row> | SpreadsheetError
          if (chunkIndex < chunks.length) {
            const chunk = chunks[chunkIndex]!
            if (offset >= chunk.length) {
              chunkIndex++
              offset = 0
              continue // chunkIndex advanced
            }
            yield* Effect.yieldNow
            const end = Math.min(chunk.length, offset + PIECE_BYTES)
            const text = decoder.push(chunk.subarray(offset, end))
            offset = end
            if (typeof text !== "string") return yield* Effect.fail(text)
            rows = parser.feed(text, false)
          } else if (!upstreamDone) {
            const next = yield* pullUpstream
            if (next === undefined) {
              upstreamDone = true
              continue // state advanced: upstreamDone
            }
            for (const c of next) inputBytes += c.length
            if (inputBytes > maxInputBytes) {
              return yield* Effect.fail(new LimitExceeded({ limit: "maxInputBytes", max: maxInputBytes, actual: inputBytes }))
            }
            chunks = next
            chunkIndex = 0
            offset = 0
            continue
          } else {
            finished = true
            const text = decoder.end()
            if (typeof text !== "string") return yield* Effect.fail(text)
            rows = parser.feed(text, true)
          }
          if (isError(rows)) return yield* Effect.fail(rows)
          if (rows.length > 0) return rows as unknown as Arr.NonEmptyReadonlyArray<Row>
        }
      })
      return pull
    }))
}

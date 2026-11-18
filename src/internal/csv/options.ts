/**
 * CSV read options, checked once when a read starts and handed to the
 * reader as a value: the reader, decoder and parser don't check them again.
 *
 * @internal
 */
import * as Result from "effect/Result"
import type { ReadOptions } from "../../csv/index.ts"
import { InvalidOption } from "../../SpreadsheetError.ts"

/** How the bytes become text: `"auto"` detects the encoding, else a decoder for the requested one. */
export type Decoding = "auto" | (() => TextDecoder)

export interface CsvOptions {
  /** The delimiter's character code, or undefined for a `sep=` line, else a guess. */
  readonly delimiter: number | undefined
  readonly decoding: Decoding
  readonly inferTypes: boolean
  readonly strict: boolean
}

/**
 * Whether `char` can delimit fields: one UTF-16 code unit other than a
 * quote, CR or LF. Used for the option and for an Excel `sep=` line.
 */
export const isDelimiter = (char: string): boolean =>
  char.length === 1 && char !== '"' && char !== "\r" && char !== "\n"

/**
 * Checks `options.delimiter` at run time, since the type can be bypassed
 * (plain JavaScript, casts): `undefined` (guess it) or a one-character
 * string other than a quote, CR or LF. Succeeds with its character code.
 */
export const parseDelimiter = (delimiter: unknown): Result.Result<number | undefined, InvalidOption> => {
  if (delimiter === undefined) return Result.succeed(undefined)
  if (typeof delimiter !== "string") {
    const type = typeof delimiter
    const kind = delimiter === null
      ? "null"
      : Array.isArray(delimiter)
      ? "an array"
      : type === "object"
      ? "an object"
      : `a ${type}`
    return Result.fail(new InvalidOption({ option: "delimiter", reason: `must be a one-character string, not ${kind}` }))
  }
  return isDelimiter(delimiter)
    ? Result.succeed(delimiter.charCodeAt(0))
    : Result.fail(
      new InvalidOption({
        option: "delimiter",
        reason: `${JSON.stringify(delimiter)} must be one character other than a quote, CR or LF`
      })
    )
}

/**
 * Checks an encoding label: `"auto"` or one `TextDecoder` knows (it only
 * says so by throwing). A known label becomes a maker of decoders for its
 * encoding: each run of the rows needs a fresh one, as decoders keep state.
 */
export const parseEncoding = (label: string): Result.Result<Decoding, InvalidOption> => {
  if (label === "auto") return Result.succeed("auto")
  let encoding: string
  try {
    encoding = new TextDecoder(label).encoding
  } catch {
    return Result.fail(
      new InvalidOption({ option: "encoding", reason: `text encoding ${JSON.stringify(label)} is not supported` })
    )
  }
  return Result.succeed(() => makeDecoder(encoding))
}

/**
 * A decoder for `encoding`, a name `TextDecoder` knows. `ignoreBOM: false`
 * makes it drop a BOM of its own encoding.
 */
export const makeDecoder = (encoding: string): TextDecoder => new TextDecoder(encoding, { ignoreBOM: false })

/**
 * The options of one read, checked: the delimiter first, then the
 * encoding; the first invalid one is the error.
 */
export const csvOptions = (options: ReadOptions | undefined): Result.Result<CsvOptions, InvalidOption> => {
  const delimiter = parseDelimiter(options?.delimiter)
  if (Result.isFailure(delimiter)) return Result.fail(delimiter.failure)
  const decoding = parseEncoding(options?.encoding ?? "auto")
  if (Result.isFailure(decoding)) return Result.fail(decoding.failure)
  return Result.succeed(withDecoding(options, delimiter.success, decoding.success))
}

/** The checked options, from the delimiter and decoding already checked. */
export const withDecoding = (
  options: ReadOptions | undefined,
  delimiter: number | undefined,
  decoding: Decoding
): CsvOptions => ({
  delimiter,
  decoding,
  inferTypes: options?.inferTypes ?? false,
  strict: options?.strict ?? false
})

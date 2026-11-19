/**
 * XML bytes -> tokenizer events. The encoding is UTF-8 unless the part
 * starts with a UTF-16 byte-order mark; invalid bytes are `MalformedFile`.
 * Multibyte characters split across chunks are carried over.
 *
 * @internal
 */
import type { Effect, Stream } from "effect"
import * as Result from "effect/Result"
import { MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import { drained, type PushDecoder, stopped } from "../decode.ts"
import type { Meter } from "../meter.ts"
import { TextInput } from "../text.ts"
import { type XmlHandler, type XmlLimits, XmlTokenizer } from "./tokenizer.ts"

/**
 * Decodes a part's bytes incrementally: UTF-8, or UTF-16 after a UTF-16
 * byte-order mark. Invalid bytes are `MalformedFile`, with the byte offset
 * in the part of the first invalid UTF-8 sequence (UTF-16 errors have no
 * offset).
 */
export const xmlText = (part: string | undefined): TextInput<MalformedFile> =>
  new TextInput({
    detect: "bom-only",
    encoding: "utf-8",
    fatal: true,
    invalid: (offset) =>
      new MalformedFile({
        format: "xml",
        reason: "invalid text encoding",
        ...(part === undefined ? {} : { part }),
        ...(offset === undefined ? {} : { offset })
      })
  })

/**
 * The push decoder of an XML part: its bytes go through the text decoder,
 * then the tokenizer, which calls `handler` and hands its work to `meter`.
 * A step's error is the first one from the decoder, the tokenizer or the
 * handler.
 */
export const xmlDecoder = (
  handler: XmlHandler,
  limits: XmlLimits,
  part: string | undefined,
  meter: Meter
): PushDecoder<never> => {
  const decoder = xmlText(part)
  const tokenizer = new XmlTokenizer(handler, limits, part, meter)
  return {
    feed: (piece) => {
      const text = decoder.push(piece)
      return stopped(Result.isSuccess(text) ? tokenizer.feed(text.success) : text.failure)
    },
    end: () => {
      const rest = decoder.end()
      return stopped(Result.isSuccess(rest) ? (tokenizer.feed(rest.success) ?? tokenizer.end()) : rest.failure)
    }
  }
}

/**
 * Tokenizes a byte stream with `handler`. Fails with the first error (from
 * the input stream, the decoder, the tokenizer or the handler). The
 * tokenizer hands its work to the work meter.
 */
export const parse = <E, R>(
  bytes: Stream.Stream<Uint8Array, E, R>,
  handler: XmlHandler,
  limits: XmlLimits,
  part?: string
): Effect.Effect<void, E | SpreadsheetError, R> => drained(bytes, (meter) => xmlDecoder(handler, limits, part, meter))

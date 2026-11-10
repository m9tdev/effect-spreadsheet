/**
 * XML bytes -> tokenizer events. The encoding is UTF-8 unless the part
 * starts with a UTF-16 byte-order mark; invalid bytes are `MalformedFile`.
 * Multibyte characters split across chunks are carried over.
 *
 * @internal
 */
import { Effect, Stream } from "effect"
import { MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import { type XmlHandler, type XmlLimits, XmlTokenizer } from "./tokenizer.ts"

/**
 * The decoder's phase. Sniffing holds the first 0-1 bytes until there are
 * enough to see a UTF-16 byte-order mark; Decoding has picked the encoding.
 */
type DecoderState =
  | { readonly _tag: "Sniffing"; readonly head: Uint8Array }
  | { readonly _tag: "Decoding"; readonly decoder: TextDecoder }

/** Decodes a part's bytes incrementally; picks the encoding from the first bytes. */
export class XmlDecoder {
  private state: DecoderState = { _tag: "Sniffing", head: new Uint8Array(0) }
  private readonly part: string | undefined

  constructor(part?: string) {
    this.part = part
  }

  private malformed(): MalformedFile {
    return new MalformedFile({
      format: "xml",
      reason: "invalid text encoding",
      ...(this.part === undefined ? {} : { part: this.part })
    })
  }

  /** Sniffing -> Decoding: picks the encoding from `head`; the new decoder, which has not read `head` yet. */
  private begin(head: Uint8Array): TextDecoder {
    const utf16 = head.length >= 2 && ((head[0] === 0xFF && head[1] === 0xFE) || (head[0] === 0xFE && head[1] === 0xFF))
    const label = !utf16 ? "utf-8" : head[0] === 0xFF ? "utf-16le" : "utf-16be"
    // ignoreBOM: false drops the byte-order mark itself
    const decoder = new TextDecoder(label, { fatal: true, ignoreBOM: false })
    this.state = { _tag: "Decoding", decoder }
    return decoder
  }

  /** Decodes the next bytes; the text available so far, or an error. */
  push(bytes: Uint8Array): string | MalformedFile {
    try {
      switch (this.state._tag) {
        case "Sniffing": {
          // need 2 bytes to see a UTF-16 byte-order mark
          const head = this.state.head.length === 0 ? bytes : concat(this.state.head, bytes)
          if (head.length < 2) {
            this.state = { _tag: "Sniffing", head }
            return ""
          }
          return this.begin(head).decode(head, { stream: true })
        }
        case "Decoding":
          return this.state.decoder.decode(bytes, { stream: true })
      }
    } catch {
      return this.malformed()
    }
  }

  /** Flushes at the end; an incomplete trailing character is an error. */
  end(): string | MalformedFile {
    try {
      switch (this.state._tag) {
        case "Sniffing": {
          const head = this.state.head
          const decoder = this.begin(head)
          return decoder.decode(head, { stream: true }) + decoder.decode()
        }
        case "Decoding":
          return this.state.decoder.decode()
      }
    } catch {
      return this.malformed()
    }
  }
}

const concat = (a: Uint8Array, b: Uint8Array): Uint8Array => {
  const out = new Uint8Array(a.length + b.length)
  out.set(a, 0)
  out.set(b, a.length)
  return out
}

/** The result of `parse`: the tokenizer's work counter, for cost checks. */
export interface ParseStats {
  readonly work: number
}

/**
 * Tokenizes a byte stream with `handler`. Fails with the first error (from
 * the input stream, the decoder, the tokenizer or the handler).
 */
export const parse = <E, R>(
  bytes: Stream.Stream<Uint8Array, E, R>,
  handler: XmlHandler,
  limits: XmlLimits,
  part?: string
): Effect.Effect<ParseStats, E | SpreadsheetError, R> =>
  Effect.suspend(() => {
    const decoder = new XmlDecoder(part)
    const tokenizer = new XmlTokenizer(handler, limits, part)
    return Stream.runForEach(bytes, (chunk) => {
      const text = decoder.push(chunk)
      if (typeof text !== "string") return Effect.fail(text)
      const error = tokenizer.feed(text)
      return error === undefined ? Effect.void : Effect.fail(error)
    }).pipe(
      Effect.flatMap(() => {
        const rest = decoder.end()
        if (typeof rest !== "string") return Effect.fail(rest)
        const error = tokenizer.feed(rest) ?? tokenizer.end()
        return error === undefined ? Effect.succeed({ work: tokenizer.work }) : Effect.fail(error)
      })
    )
  })

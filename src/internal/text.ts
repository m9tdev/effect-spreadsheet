/**
 * Bytes -> text for the text formats (CSV, XML): a UTF-8 validator, the
 * encoding detection, and one incremental decoder that sniffs a byte-order
 * mark (and, for CSV, the encoding) before decoding. Multibyte characters
 * split across chunks are carried over by `TextDecoder`'s `stream` mode.
 *
 * @internal
 */
import { absurd } from "effect/Function"
import * as Result from "effect/Result"

/**
 * The byte offset of the first ill-formed UTF-8 sequence in `bytes` (the
 * offset of its first byte), or -1 if there is none. Ill-formed is what a
 * `fatal` `TextDecoder` rejects: a byte that can't start a sequence, a
 * lead byte not followed by the continuation bytes it needs, overlong
 * forms, surrogates and code points past U+10FFFF. `complete` says
 * `bytes` is the whole input, so a sequence cut short at the end is
 * ill-formed; otherwise it may go on in later bytes, and is accepted if
 * its bytes so far can start a sequence.
 */
export const firstInvalidUtf8 = (bytes: Uint8Array, complete: boolean): number => {
  const n = bytes.length
  let i = 0
  // progress: every branch either returns or moves i forward by 1 to 4
  while (i < n) {
    const b = bytes[i]!
    if (b < 0x80) {
      i++
      continue
    }
    // continuation bytes needed, and the range of the first one
    let need: number
    let lo = 0x80
    let hi = 0xBF
    if (b >= 0xC2 && b <= 0xDF) need = 1
    else if (b >= 0xE0 && b <= 0xEF) {
      need = 2
      if (b === 0xE0) lo = 0xA0 // overlong
      else if (b === 0xED) hi = 0x9F // surrogates
    } else if (b >= 0xF0 && b <= 0xF4) {
      need = 3
      if (b === 0xF0) lo = 0x90 // overlong
      else if (b === 0xF4) hi = 0x8F // past U+10FFFF
    } else return i
    // progress: k goes from 1 up to need (at most 3)
    for (let k = 1; k <= need; k++) {
      if (i + k >= n) return complete ? i : -1
      const c = bytes[i + k]!
      if (k === 1 ? c < lo || c > hi : c < 0x80 || c > 0xBF) return i
    }
    i += need + 1
  }
  return -1
}

/** Bytes buffered before deciding the encoding in `"bom-utf16-utf8-1252"` mode. */
export const DETECT_BYTES = 64 * 1024

/** The encoding a byte-order mark names, if `head` starts with one. */
export const bomEncoding = (head: Uint8Array): string | undefined => {
  const n = head.length
  if (n >= 3 && head[0] === 0xEF && head[1] === 0xBB && head[2] === 0xBF) return "utf-8"
  if (n >= 2 && head[0] === 0xFF && head[1] === 0xFE) return "utf-16le"
  if (n >= 2 && head[0] === 0xFE && head[1] === 0xFF) return "utf-16be"
  return undefined
}

/**
 * Picks an encoding from the first bytes: the one a byte-order mark names,
 * else UTF-16 if every other byte is mostly NUL, else UTF-8 if the bytes
 * are valid UTF-8, else windows-1252. `complete` says `head` is the whole
 * input (so a multibyte sequence cut at its end is invalid).
 */
export const detectEncoding = (head: Uint8Array, complete: boolean): string => {
  const n = head.length
  const bom = bomEncoding(head)
  if (bom !== undefined) return bom
  // UTF-16 without a BOM: mostly-ASCII text has a NUL in every other byte.
  const sample = Math.min(n, 4096) & ~1
  if (sample >= 4) {
    let evenZeros = 0
    let oddZeros = 0
    // progress: i += 2 up to `sample`
    for (let i = 0; i < sample; i += 2) {
      if (head[i] === 0) evenZeros++
      if (head[i + 1] === 0) oddZeros++
    }
    const pairs = sample / 2
    if (oddZeros >= pairs * 0.4 && evenZeros <= pairs * 0.05) return "utf-16le"
    if (evenZeros >= pairs * 0.4 && oddZeros <= pairs * 0.05) return "utf-16be"
  }
  return firstInvalidUtf8(head, complete) < 0 ? "utf-8" : "windows-1252"
}

/**
 * How the encoding is chosen. A byte-order mark always wins (it is
 * dropped, and its encoding is used). Without one:
 * - `"bom-only"`: `encoding`;
 * - `"bom-utf16-utf8-1252"`: detected from the first {@link DETECT_BYTES}
 *   bytes (or the whole input, if shorter) by {@link detectEncoding}.
 */
export type Detect =
  | { readonly detect: "bom-only"; readonly encoding: string }
  | { readonly detect: "bom-utf16-utf8-1252" }

/**
 * What invalid bytes do: with `fatal: false` they become U+FFFD; with
 * `fatal: true` they are the error `invalid` makes, given the byte offset
 * of the first ill-formed UTF-8 sequence (undefined in other encodings,
 * which the platform validates without saying where).
 */
export type OnInvalid<E> =
  | { readonly fatal: false }
  | { readonly fatal: true; readonly invalid: (offset: number | undefined) => E }

/**
 * How decoded bytes are checked: not at all (`fatal: false`); for UTF-8,
 * by looking for U+FFFD in the text and, only then, by the validator (the
 * input may hold U+FFFD itself); for other encodings, by the platform's
 * fatal decoder, which only says so by throwing.
 */
type Check<E> =
  | { readonly _tag: "None" }
  | { readonly _tag: "Utf8"; readonly invalid: (offset: number | undefined) => E }
  | { readonly _tag: "Platform"; readonly invalid: (offset: number | undefined) => E }

/**
 * Where a {@link TextInput} is: holding the first bytes back until the
 * encoding is decided, or decoding with the chosen encoding.
 */
type Phase<E> =
  | { readonly _tag: "Detecting"; readonly head: Array<Uint8Array>; bytes: number }
  | Decoding<E>

interface Decoding<E> {
  readonly _tag: "Decoding"
  readonly decoder: TextDecoder
  readonly check: Check<E>
}

const NOTHING: Result.Result<string, never> = Result.succeed("")

/**
 * Feeds bytes in, gets text out, incrementally. The first bytes are held
 * back until the encoding is decided: {@link DETECT_BYTES} of them (or the
 * whole input, if shorter) when detecting, else 3, the longest byte-order
 * mark. So the decoder's first call always sees a whole mark: `TextDecoder`
 * implementations differ on a mark split across calls (some drop a second
 * mark after it too).
 */
export class TextInput<E = never> {
  private phase: Phase<E> = { _tag: "Detecting", head: [], bytes: 0 }
  private readonly detect: Detect
  private readonly onInvalid: OnInvalid<E>
  private readonly hold: number
  /** Bytes decoded so far (after the held ones are released). */
  private consumed = 0
  /** The last (at most 3) of those bytes: where a UTF-8 sequence cut at a chunk's end starts. */
  private tail: Uint8Array = new Uint8Array(0)

  constructor(options: Detect & OnInvalid<E>) {
    this.detect = options.detect === "bom-only"
      ? { detect: "bom-only", encoding: options.encoding }
      : { detect: "bom-utf16-utf8-1252" }
    this.onInvalid = options.fatal ? { fatal: true, invalid: options.invalid } : { fatal: false }
    this.hold = this.detect.detect === "bom-utf16-utf8-1252" ? DETECT_BYTES : 3
  }

  /** Decodes a chunk: the text available so far, or the error for invalid bytes. */
  push(bytes: Uint8Array): Result.Result<string, E> {
    const phase = this.phase
    switch (phase._tag) {
      case "Detecting": {
        phase.head.push(bytes)
        phase.bytes += bytes.length
        if (phase.bytes < this.hold) return NOTHING
        const head = concat(phase.head, phase.bytes)
        return this.decode(this.begin(head, false), head, false)
      }
      case "Decoding":
        return this.decode(phase, bytes, false)
      default:
        return absurd(phase)
    }
  }

  /** Flushes at the end of the input: an incomplete trailing sequence is U+FFFD, or the error. */
  end(): Result.Result<string, E> {
    const phase = this.phase
    switch (phase._tag) {
      case "Detecting": {
        const head = concat(phase.head, phase.bytes)
        return this.decode(this.begin(head, true), head, true)
      }
      case "Decoding":
        return this.decode(phase, new Uint8Array(0), true)
      default:
        return absurd(phase)
    }
  }

  /**
   * Leaves detection: picks the decoder for the held-back bytes `head`
   * (`complete` says it is the whole input).
   */
  private begin(head: Uint8Array, complete: boolean): Decoding<E> {
    const detect = this.detect
    const encoding = bomEncoding(head) ?? (
      detect.detect === "bom-only"
        ? detect.encoding
        // detect on exactly the first DETECT_BYTES, so the choice doesn't
        // depend on how the input was chunked
        : head.length > DETECT_BYTES
        ? detectEncoding(head.subarray(0, DETECT_BYTES), false)
        : detectEncoding(head, complete)
    )
    const onInvalid = this.onInvalid
    const check: Check<E> = !onInvalid.fatal
      ? { _tag: "None" }
      : { _tag: encoding === "utf-8" ? "Utf8" : "Platform", invalid: onInvalid.invalid }
    // ignoreBOM: false drops a byte-order mark of the decoder's own encoding
    const decoder = new TextDecoder(encoding, { fatal: check._tag === "Platform", ignoreBOM: false })
    const phase: Decoding<E> = { _tag: "Decoding", decoder, check }
    this.phase = phase
    return phase
  }

  /** Decodes `bytes` (and flushes, if `last`), checked as the phase says. */
  private decode(
    phase: Decoding<E>,
    bytes: Uint8Array,
    last: boolean
  ): Result.Result<string, E> {
    const { check, decoder } = phase
    switch (check._tag) {
      case "None":
        return Result.succeed(decodeAll(decoder, bytes, last))
      case "Platform":
        try {
          return Result.succeed(decodeAll(decoder, bytes, last))
        } catch {
          return Result.fail(check.invalid(undefined))
        }
      case "Utf8": {
        const text = decodeAll(decoder, bytes, last)
        // U+FFFD stands for invalid bytes, unless the input holds it itself
        if (text.indexOf("\uFFFD") >= 0) {
          const at = this.invalidAt(bytes, last)
          if (at >= 0) return Result.fail(check.invalid(at))
        }
        this.advance(bytes)
        return Result.succeed(text)
      }
      default:
        return absurd(check)
    }
  }

  /**
   * The byte offset in the input of the first ill-formed UTF-8 sequence in
   * the bytes before `bytes` and `bytes` itself, or -1. The bytes before
   * were valid as far as they went, so only a sequence they left
   * unfinished can go wrong: it starts in the last 3 bytes (`tail`), at
   * its first byte that isn't a continuation byte.
   */
  private invalidAt(bytes: Uint8Array, complete: boolean): number {
    const tail = this.tail
    let from = 0
    // progress: from increases up to tail.length (at most 3)
    while (from < tail.length && (tail[from]! & 0xC0) === 0x80) from++
    const carried = tail.subarray(from)
    const at = firstInvalidUtf8(carried.length === 0 ? bytes : concat([carried, bytes], carried.length + bytes.length), complete)
    return at < 0 ? -1 : this.consumed - carried.length + at
  }

  /** Counts `bytes` as decoded, and keeps the last 3 decoded bytes. */
  private advance(bytes: Uint8Array): void {
    const n = bytes.length
    this.consumed += n
    if (n >= 3) this.tail = bytes.subarray(n - 3)
    else if (n > 0) {
      const joined = concat([this.tail, bytes], this.tail.length + n)
      this.tail = joined.subarray(Math.max(0, joined.length - 3))
    }
  }
}

/** `bytes` decoded in stream mode, then the decoder flushed if `last`. */
const decodeAll = (decoder: TextDecoder, bytes: Uint8Array, last: boolean): string => {
  const text = bytes.length === 0 ? "" : decoder.decode(bytes, { stream: true })
  return last ? text + decoder.decode() : text
}

const concat = (parts: ReadonlyArray<Uint8Array>, total: number): Uint8Array => {
  if (parts.length === 1) return parts[0]!
  const out = new Uint8Array(total)
  let offset = 0
  // progress: one part per iteration
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

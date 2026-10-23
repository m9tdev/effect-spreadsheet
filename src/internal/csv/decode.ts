/**
 * Incremental bytes -> text decoding for delimited text, with BOM and
 * encoding detection. Multibyte characters split across chunks are carried
 * over by `TextDecoder`'s `stream` mode.
 *
 * @internal
 */
import { UnsupportedFormat } from "../../SpreadsheetError.ts"

/** Bytes buffered before deciding the encoding in `"auto"` mode. */
export const DETECT_BYTES = 64 * 1024

/**
 * Picks an encoding from the first bytes. `complete` says `head` is the
 * whole input (so a multibyte sequence cut at its end is invalid).
 */
export const detectEncoding = (head: Uint8Array, complete: boolean): string => {
  const n = head.length
  if (n >= 3 && head[0] === 0xEF && head[1] === 0xBB && head[2] === 0xBF) return "utf-8"
  if (n >= 2 && head[0] === 0xFF && head[1] === 0xFE) return "utf-16le"
  if (n >= 2 && head[0] === 0xFE && head[1] === 0xFF) return "utf-16be"
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
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(head, { stream: !complete })
    return "utf-8"
  } catch {
    return "windows-1252"
  }
}

const makeDecoder = (label: string): TextDecoder | UnsupportedFormat => {
  try {
    // ignoreBOM: false makes the decoder drop a BOM of its own encoding.
    return new TextDecoder(label, { ignoreBOM: false })
  } catch {
    return new UnsupportedFormat({ format: "csv", reason: `text encoding ${JSON.stringify(label)} is not supported` })
  }
}

/**
 * Feeds bytes in, gets text out. With `encoding: "auto"` the first
 * {@link DETECT_BYTES} bytes (or the whole input, if shorter) are held back
 * until the encoding is decided.
 */
export class ByteDecoder {
  private decoder: TextDecoder | undefined
  private head: Array<Uint8Array> = []
  private headBytes = 0
  /** The encoding in use, once decided. */
  encoding: string | undefined

  private readonly requested: string

  constructor(requested: string) {
    this.requested = requested
  }

  /** Decodes a chunk. Returns the text available so far, or an error. */
  push(bytes: Uint8Array): string | UnsupportedFormat {
    if (this.decoder === undefined) {
      if (this.requested === "auto") {
        this.head.push(bytes)
        this.headBytes += bytes.length
        if (this.headBytes < DETECT_BYTES) return ""
        return this.start(false)
      }
      const decoder = makeDecoder(this.requested)
      if (decoder instanceof UnsupportedFormat) return decoder
      this.decoder = decoder
      this.encoding = this.requested
    }
    return this.decoder.decode(bytes, { stream: true })
  }

  /** Flushes at end of input (an incomplete trailing sequence becomes U+FFFD). */
  end(): string | UnsupportedFormat {
    let text = ""
    if (this.decoder === undefined) {
      const started = this.start(true)
      if (typeof started !== "string") return started
      text = started
    }
    return text + this.decoder!.decode()
  }

  private start(complete: boolean): string | UnsupportedFormat {
    const head = concat(this.head, this.headBytes)
    this.head = []
    // Detect on exactly the first DETECT_BYTES, so the choice doesn't depend
    // on how the input was chunked.
    const label = this.requested === "auto"
      ? head.length > DETECT_BYTES
        ? detectEncoding(head.subarray(0, DETECT_BYTES), false)
        : detectEncoding(head, complete)
      : this.requested
    const decoder = makeDecoder(label)
    if (decoder instanceof UnsupportedFormat) return decoder
    this.decoder = decoder
    this.encoding = label
    return decoder.decode(head, { stream: true })
  }
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

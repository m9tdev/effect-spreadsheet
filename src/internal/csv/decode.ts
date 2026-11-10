/**
 * Incremental bytes -> text decoding for delimited text, with BOM and
 * encoding detection. Multibyte characters split across chunks are carried
 * over by `TextDecoder`'s `stream` mode.
 *
 * @internal
 */
import { absurd } from "effect/Function"
import { InvalidOption } from "../../SpreadsheetError.ts"

/** Bytes buffered before deciding the encoding in `"auto"` mode. */
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
 * Picks an encoding from the first bytes. `complete` says `head` is the
 * whole input (so a multibyte sequence cut at its end is invalid).
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
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(head, { stream: !complete })
    return "utf-8"
  } catch {
    return "windows-1252"
  }
}

/** An error if `label` isn't `"auto"` or a label `TextDecoder` knows. */
export const validateEncoding = (label: string): InvalidOption | undefined => {
  if (label === "auto") return undefined
  const decoder = makeDecoder(label)
  return decoder instanceof InvalidOption ? decoder : undefined
}

const makeDecoder = (label: string): TextDecoder | InvalidOption => {
  try {
    // ignoreBOM: false makes the decoder drop a BOM of its own encoding.
    return new TextDecoder(label, { ignoreBOM: false })
  } catch {
    return new InvalidOption({ option: "encoding", reason: `text encoding ${JSON.stringify(label)} is not supported` })
  }
}

/**
 * Where a {@link ByteDecoder} is: holding the first bytes back until the
 * encoding is decided, decoding with the chosen encoding, or stopped on an
 * error (which every later call returns again).
 */
type Phase =
  | { readonly _tag: "Detecting"; readonly head: Array<Uint8Array>; bytes: number }
  | Decided

type Decided =
  | { readonly _tag: "Decoding"; readonly decoder: TextDecoder }
  | { readonly _tag: "Failed"; readonly error: InvalidOption }

/**
 * Leaves detection: picks the encoding from the held-back bytes `head`
 * (`complete` says it is the whole input) and makes its decoder.
 */
const decide = (requested: string, head: Uint8Array, complete: boolean): Decided => {
  if (requested !== "auto") {
    const invalid = validateEncoding(requested)
    if (invalid !== undefined) return { _tag: "Failed", error: invalid }
  }
  // Detect on exactly the first DETECT_BYTES, so the choice doesn't depend
  // on how the input was chunked.
  const label = bomEncoding(head) ?? (requested === "auto"
    ? head.length > DETECT_BYTES
      ? detectEncoding(head.subarray(0, DETECT_BYTES), false)
      : detectEncoding(head, complete)
    : requested)
  const decoder = makeDecoder(label)
  if (decoder instanceof InvalidOption) return { _tag: "Failed", error: decoder }
  return { _tag: "Decoding", decoder }
}

/**
 * Feeds bytes in, gets text out. A BOM always wins (it is dropped, and its
 * encoding is used even when another one is requested). With `"auto"` the
 * first {@link DETECT_BYTES} bytes (or the whole input, if shorter) are held
 * back until the encoding is decided; with a requested encoding only the
 * first 3 bytes are, to look for a BOM.
 */
export class ByteDecoder {
  private phase: Phase = { _tag: "Detecting", head: [], bytes: 0 }
  private readonly requested: string

  constructor(requested: string) {
    this.requested = requested
  }

  /** Decodes a chunk. Returns the text available so far, or an error. */
  push(bytes: Uint8Array): string | InvalidOption {
    const phase = this.phase
    switch (phase._tag) {
      case "Detecting": {
        phase.head.push(bytes)
        phase.bytes += bytes.length
        if (phase.bytes < (this.requested === "auto" ? DETECT_BYTES : 3)) return ""
        const head = concat(phase.head, phase.bytes)
        const next = this.phase = decide(this.requested, head, false)
        return next._tag === "Decoding" ? next.decoder.decode(head, { stream: true }) : next.error
      }
      case "Decoding":
        return phase.decoder.decode(bytes, { stream: true })
      case "Failed":
        return phase.error
      default:
        return absurd(phase)
    }
  }

  /** Flushes at end of input (an incomplete trailing sequence becomes U+FFFD). */
  end(): string | InvalidOption {
    const phase = this.phase
    switch (phase._tag) {
      case "Detecting": {
        const head = concat(phase.head, phase.bytes)
        const next = this.phase = decide(this.requested, head, true)
        return next._tag === "Decoding"
          ? next.decoder.decode(head, { stream: true }) + next.decoder.decode()
          : next.error
      }
      case "Decoding":
        return phase.decoder.decode()
      case "Failed":
        return phase.error
      default:
        return absurd(phase)
    }
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

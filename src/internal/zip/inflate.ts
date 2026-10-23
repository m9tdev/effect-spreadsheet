/**
 * Raw-deflate decompression behind one small interface, so another
 * implementation (e.g. a TypeScript inflate) can replace the default one.
 * The default uses the web-standard `DecompressionStream("deflate-raw")`.
 *
 * @internal
 */

/** Pull-based decompression of one entry. */
export interface Inflater {
  /**
   * The next chunk of decompressed bytes, or `undefined` at the end. Rejects
   * on invalid, truncated or trailing data.
   */
  readonly next: () => Promise<Uint8Array | undefined>
  /** Stops decompressing and releases the resources (idempotent). */
  readonly cancel: () => void
}

/** Creates an `Inflater` for one compressed buffer. */
export type InflateRaw = (compressed: Uint8Array) => Inflater

/** Compressed bytes are handed to the decompressor in slices of this size. */
const FEED_BYTES = 64 * 1024

/**
 * The default implementation. Compressed input is written in slices while
 * output is read; with stream backpressure, the decompressor never runs far
 * ahead of the reader, so a highly compressed entry is only inflated as far
 * as it is read.
 */
export const inflateRawWeb: InflateRaw = (compressed) => {
  const stream = new DecompressionStream("deflate-raw")
  const writer = stream.writable.getWriter()
  const reader = stream.readable.getReader()
  let cancelled = false
  const feed = async () => {
    // progress: at increases by FEED_BYTES up to compressed.length
    for (let at = 0; at < compressed.length && !cancelled; at += FEED_BYTES) {
      await writer.write(compressed.subarray(at, Math.min(compressed.length, at + FEED_BYTES)) as Uint8Array<ArrayBuffer>)
    }
    if (!cancelled) await writer.close()
  }
  // Errors surface on the read side; this only keeps them from being unhandled.
  feed().catch(() => {})
  return {
    next: async () => {
      const result = await reader.read()
      return result.done ? undefined : result.value
    },
    cancel: () => {
      if (cancelled) return
      cancelled = true
      reader.cancel().catch(() => {})
      writer.abort().catch(() => {})
    }
  }
}

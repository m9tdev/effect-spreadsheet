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
   * on invalid, truncated or trailing data, after the bytes decompressed
   * before the error have been returned (see `inflateRawWeb` for its limit).
   */
  readonly next: () => Promise<Uint8Array | undefined>
  /** Stops decompressing and releases the resources (idempotent). */
  readonly cancel: () => void
}

/** Creates an `Inflater` for one compressed buffer. */
export type InflateRaw = (compressed: Uint8Array) => Inflater

/**
 * Compressed bytes are handed to the decompressor in slices of this size;
 * the next slice only once everything decompressed so far has been handed
 * out. Not smaller than Node's 64 KiB write buffer: a slice's write then
 * settles only once the slice is decompressed, so the writer is closed
 * after the last slice is processed, never while it waits (Node would
 * decompress a waiting last slice together with the final flush, and drop
 * its output if that fails).
 */
const FEED_BYTES = 64 * 1024

/**
 * Most decompressed bytes taken off the stream and held ahead of the
 * reader. Past it, reading pauses until the reader takes some, so a slice
 * that inflates hugely (a deflate bomb) is never held whole in memory.
 */
export const HOLD_BYTES = 4 * 1024 * 1024

/**
 * The default implementation. A read is kept pending on the stream, so its
 * output is taken as soon as it is produced: a stream that errors drops the
 * chunks still queued in it, and which ones are queued depends on timing,
 * so without this the bytes seen before an error (and so which error a
 * reader reports first) would change from run to run. Input is written
 * slice by slice, the next one only when the output taken so far has been
 * handed out, so an entry is only inflated about as far as it is read.
 *
 * Known limit: while `HOLD_BYTES` are held no read is pending, so for a
 * damaged entry whose 64 KiB slice inflates past `HOLD_BYTES`, the output
 * kept before its error can still depend on timing.
 */
export const inflateRawWeb: InflateRaw = (compressed) => {
  const stream = new DecompressionStream("deflate-raw")
  const writer = stream.writable.getWriter()
  const reader = stream.readable.getReader()
  /** Output taken off the stream, not handed out yet, and its size. */
  const ready: Array<Uint8Array> = []
  let held = 0
  let state: { readonly _tag: "Open" } | { readonly _tag: "Ended" } | { readonly _tag: "Failed"; readonly error: unknown } = {
    _tag: "Open"
  }
  /** Compressed bytes written so far; past the end once the writer is closed. */
  let at = 0
  let writing = false
  let cancelled = false
  /** Wakes `next` when a chunk arrives, a write settles or the stream ends. */
  let wakeNext: (() => void) | undefined
  /** Wakes the pump when it paused at `HOLD_BYTES` and some output was taken. */
  let wakePump: (() => void) | undefined
  const notifyNext = () => {
    const wake = wakeNext
    wakeNext = undefined
    wake?.()
  }
  const notifyPump = () => {
    const wake = wakePump
    wakePump = undefined
    wake?.()
  }
  const pump = async () => {
    try {
      // progress: each read takes a chunk or ends (done or error); each
      // pause ends once the reader took a chunk (held decreases) or cancelled
      while (!cancelled) {
        if (held >= HOLD_BYTES) {
          await new Promise<void>((resolve) => {
            wakePump = resolve
          })
          continue
        }
        const result = await reader.read()
        if (result.done) break
        ready.push(result.value)
        held += result.value.length
        notifyNext()
      }
      state = { _tag: "Ended" }
    } catch (error) {
      state = { _tag: "Failed", error }
    }
    notifyNext()
  }
  void pump()
  /** Writes the next slice, or closes the writer after the last one. Errors surface on the read side. */
  const feed = () => {
    writing = true
    const settled = () => {
      writing = false
      notifyNext()
    }
    if (at < compressed.length) {
      const slice = compressed.subarray(at, Math.min(compressed.length, at + FEED_BYTES))
      at += slice.length
      writer.write(slice as Uint8Array<ArrayBuffer>).then(settled, () => {}) // a failed write leaves `writing` set: no more input
    } else {
      at = compressed.length + 1
      writer.close().then(settled, () => {})
    }
  }
  return {
    next: async () => {
      // progress: each wait ends with a chunk taken, a slice written, or the end
      while (true) {
        const chunk = ready.shift()
        if (chunk !== undefined) {
          held -= chunk.length
          if (held < HOLD_BYTES) notifyPump()
          return chunk
        }
        if (state._tag === "Failed") throw state.error
        if (state._tag === "Ended" || cancelled) return undefined
        if (!writing && at <= compressed.length) feed()
        await new Promise<void>((resolve) => {
          wakeNext = resolve
        })
      }
    },
    cancel: () => {
      if (cancelled) return
      cancelled = true
      reader.cancel().catch(() => {})
      writer.abort().catch(() => {})
      notifyNext()
      notifyPump()
    }
  }
}

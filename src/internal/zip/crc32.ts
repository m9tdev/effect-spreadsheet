/**
 * CRC-32 (IEEE 802.3, reflected polynomial 0xEDB88320), as used by zip.
 *
 * @internal
 */

const TABLE = (() => {
  const table = new Uint32Array(256)
  // progress: n increases up to 256; k up to 8
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

/** The initial running value. */
export const CRC32_START = 0xFFFFFFFF

/** Updates a running value with `bytes`. */
export const crc32Update = (running: number, bytes: Uint8Array): number => {
  let c = running
  // progress: i increases up to bytes.length
  for (let i = 0; i < bytes.length; i++) c = TABLE[(c ^ bytes[i]!) & 0xFF]! ^ (c >>> 8)
  return c >>> 0
}

/** The final CRC from a running value. */
export const crc32Final = (running: number): number => (running ^ 0xFFFFFFFF) >>> 0

/** CRC-32 of `bytes`. */
export const crc32 = (bytes: Uint8Array): number => crc32Final(crc32Update(CRC32_START, bytes))

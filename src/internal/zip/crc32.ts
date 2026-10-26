/**
 * CRC-32 (IEEE 802.3, reflected polynomial 0xEDB88320), as used by zip.
 *
 * @internal
 */

/**
 * Lookup tables for "slicing by 4": TABLES[0] is the classic byte table,
 * TABLES[k][n] is the CRC of byte n followed by k zero bytes, so four
 * bytes are folded per step.
 */
const TABLES = (() => {
  const tables = Array.from({ length: 4 }, () => new Uint32Array(256))
  const t0 = tables[0]!
  // progress: n increases up to 256; k up to 8
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1
    t0[n] = c >>> 0
  }
  // progress: k up to 4, n up to 256
  for (let k = 1; k < 4; k++) {
    const prev = tables[k - 1]!
    const table = tables[k]!
    for (let n = 0; n < 256; n++) table[n] = (t0[prev[n]! & 0xFF]! ^ (prev[n]! >>> 8)) >>> 0
  }
  return tables
})()
const [T0, T1, T2, T3] = TABLES as [Uint32Array, Uint32Array, Uint32Array, Uint32Array]

/** The initial running value. */
export const CRC32_START = 0xFFFFFFFF

/** Updates a running value with `bytes`. */
export const crc32Update = (running: number, bytes: Uint8Array): number => {
  let c = running
  const n = bytes.length
  let i = 0
  // progress: i increases by 4 up to n - 3
  for (; i + 4 <= n; i += 4) {
    c ^= bytes[i]! | (bytes[i + 1]! << 8) | (bytes[i + 2]! << 16) | (bytes[i + 3]! << 24)
    c = T3[c & 0xFF]! ^ T2[(c >>> 8) & 0xFF]! ^ T1[(c >>> 16) & 0xFF]! ^ T0[c >>> 24]!
  }
  // progress: i increases up to n
  for (; i < n; i++) c = T0[(c ^ bytes[i]!) & 0xFF]! ^ (c >>> 8)
  return c >>> 0
}

/** The final CRC from a running value. */
export const crc32Final = (running: number): number => (running ^ 0xFFFFFFFF) >>> 0

/** CRC-32 of `bytes`. */
export const crc32 = (bytes: Uint8Array): number => crc32Final(crc32Update(CRC32_START, bytes))

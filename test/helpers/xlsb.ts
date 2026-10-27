/** BIFF12 record writers for tests (synthetic content only). */

/** A BIFF12 variable-length integer: 7 bits per byte, low bits first. */
export const varint = (n: number): Array<number> => {
  const out: Array<number> = []
  let rest = n
  do {
    const low = rest % 128
    rest = Math.floor(rest / 128)
    out.push(rest > 0 ? low | 0x80 : low)
  } while (rest > 0)
  return out
}

export const u8 = (n: number) => [n & 0xFF]
export const u16 = (n: number) => [n & 0xFF, (n >>> 8) & 0xFF]
export const u32 = (n: number) => [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF]
export const f64 = (n: number) => {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return [...b]
}

/** An XLWideString: u32 count, UTF-16LE code units. */
export const wide = (s: string): Array<number> => {
  const out = u32(s.length)
  // progress: k up to s.length
  for (let k = 0; k < s.length; k++) out.push(...u16(s.charCodeAt(k)))
  return out
}

/** An XLNullableWideString; `undefined` is the null string. */
export const nullableWide = (s: string | undefined): Array<number> => (s === undefined ? u32(0xFFFFFFFF) : wide(s))

/** One record: type and size varints, then the data. */
export const record = (type: number, data: ReadonlyArray<number> | Uint8Array = []): Uint8Array =>
  new Uint8Array([...varint(type), ...varint(data.length), ...data])

export const concat = (parts: ReadonlyArray<Uint8Array>): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

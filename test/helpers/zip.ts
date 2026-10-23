/** Minimal stored-only zip writer for tests (CRCs are zero: only the directory is read). */
export const makeZip = (entries: ReadonlyArray<{ name: string; data?: Uint8Array }>): Uint8Array => {
  const te = new TextEncoder()
  const locals: Array<Uint8Array> = []
  const centrals: Array<Uint8Array> = []
  let offset = 0
  for (const entry of entries) {
    const name = te.encode(entry.name)
    const data = entry.data ?? new Uint8Array(0)
    const local = new Uint8Array(30 + name.length + data.length)
    const lv = new DataView(local.buffer)
    lv.setUint32(0, 0x04034b50, true)
    lv.setUint16(4, 20, true)
    lv.setUint32(18, data.length, true)
    lv.setUint32(22, data.length, true)
    lv.setUint16(26, name.length, true)
    local.set(name, 30)
    local.set(data, 30 + name.length)
    const central = new Uint8Array(46 + name.length)
    const cv = new DataView(central.buffer)
    cv.setUint32(0, 0x02014b50, true)
    cv.setUint16(4, 20, true)
    cv.setUint16(6, 20, true)
    cv.setUint32(20, data.length, true)
    cv.setUint32(24, data.length, true)
    cv.setUint16(28, name.length, true)
    cv.setUint32(42, offset, true)
    central.set(name, 46)
    locals.push(local)
    centrals.push(central)
    offset += local.length
  }
  const cdSize = centrals.reduce((a, c) => a + c.length, 0)
  const eocd = new Uint8Array(22)
  const ev = new DataView(eocd.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, entries.length, true)
  ev.setUint16(10, entries.length, true)
  ev.setUint32(12, cdSize, true)
  ev.setUint32(16, offset, true)
  const out = new Uint8Array(offset + cdSize + 22)
  let p = 0
  for (const part of [...locals, ...centrals, eocd]) {
    out.set(part, p)
    p += part.length
  }
  return out
}

/** Raw-deflates `data` with the web-standard CompressionStream (tests only). */
export const deflateRaw = async (data: Uint8Array): Promise<Uint8Array> => {
  const stream = new CompressionStream("deflate-raw")
  const writer = stream.writable.getWriter()
  void writer.write(data as Uint8Array<ArrayBuffer>)
  void writer.close()
  const parts: Array<Uint8Array> = []
  const reader = stream.readable.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    parts.push(value)
  }
  return concat(parts)
}

export const concat = (parts: ReadonlyArray<Uint8Array>): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1
    t[n] = c >>> 0
  }
  return t
})()

/** Independent CRC-32 for test fixtures. */
export const testCrc32 = (data: Uint8Array): number => {
  let c = 0xFFFFFFFF
  for (const b of data) c = CRC_TABLE[(c ^ b) & 0xFF]! ^ (c >>> 8)
  return (c ^ 0xFFFFFFFF) >>> 0
}

export interface BuildEntry {
  readonly name: string
  readonly data: Uint8Array
  readonly method?: "store" | "deflate"
  /** Overrides written into the central directory (and local header) for hostile cases. */
  readonly crc32?: number
  readonly uncompressedSize?: number
  readonly compressedSize?: number
  readonly flags?: number
  readonly methodCode?: number
  /** Write 0xFFFFFFFF sizes/offset in the central directory plus a ZIP64 extra field. */
  readonly zip64?: boolean
  /** Bytes appended after the compressed data (inside the entry's declared size). */
  readonly trailing?: Uint8Array
}

/** A zip with real CRCs and optional deflate (tests only). */
export const buildZip = async (entries: ReadonlyArray<BuildEntry>): Promise<Uint8Array> => {
  const te = new TextEncoder()
  const locals: Array<Uint8Array> = []
  const centrals: Array<Uint8Array> = []
  let offset = 0
  for (const entry of entries) {
    const name = te.encode(entry.name)
    const deflate = (entry.method ?? "deflate") === "deflate"
    let body = deflate ? await deflateRaw(entry.data) : entry.data
    if (entry.trailing !== undefined) body = concat([body, entry.trailing])
    const crc = entry.crc32 ?? testCrc32(entry.data)
    const usize = entry.uncompressedSize ?? entry.data.length
    const csize = entry.compressedSize ?? body.length
    const method = entry.methodCode ?? (deflate ? 8 : 0)
    const flags = entry.flags ?? 0x800
    const local = new Uint8Array(30 + name.length + body.length)
    const lv = new DataView(local.buffer)
    lv.setUint32(0, 0x04034b50, true)
    lv.setUint16(4, 20, true)
    lv.setUint16(6, flags, true)
    lv.setUint16(8, method, true)
    lv.setUint32(14, crc, true)
    lv.setUint32(18, csize, true)
    lv.setUint32(22, usize, true)
    lv.setUint16(26, name.length, true)
    local.set(name, 30)
    local.set(body, 30 + name.length)
    const extra = entry.zip64 ? new Uint8Array(4 + 24) : new Uint8Array(0)
    if (entry.zip64) {
      const ev = new DataView(extra.buffer)
      ev.setUint16(0, 0x0001, true)
      ev.setUint16(2, 24, true)
      ev.setBigUint64(4, BigInt(usize), true)
      ev.setBigUint64(12, BigInt(csize), true)
      ev.setBigUint64(20, BigInt(offset), true)
    }
    const central = new Uint8Array(46 + name.length + extra.length)
    const cv = new DataView(central.buffer)
    cv.setUint32(0, 0x02014b50, true)
    cv.setUint16(4, 45, true)
    cv.setUint16(6, 20, true)
    cv.setUint16(8, flags, true)
    cv.setUint16(10, method, true)
    cv.setUint32(16, crc, true)
    cv.setUint32(20, entry.zip64 ? 0xFFFFFFFF : csize, true)
    cv.setUint32(24, entry.zip64 ? 0xFFFFFFFF : usize, true)
    cv.setUint16(28, name.length, true)
    cv.setUint16(30, extra.length, true)
    cv.setUint32(42, entry.zip64 ? 0xFFFFFFFF : offset, true)
    central.set(name, 46)
    central.set(extra, 46 + name.length)
    locals.push(local)
    centrals.push(central)
    offset += local.length
  }
  const cdSize = centrals.reduce((a, c) => a + c.length, 0)
  const eocd = new Uint8Array(22)
  const ev = new DataView(eocd.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, entries.length, true)
  ev.setUint16(10, entries.length, true)
  ev.setUint32(12, cdSize, true)
  ev.setUint32(16, offset, true)
  return concat([...locals, ...centrals, eocd])
}

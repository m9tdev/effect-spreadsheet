/**
 * A small CFB (version 3, 512-byte sectors) writer for tests: streams and
 * storages, small streams in the mini stream, a DIFAT when the FAT needs
 * more than 109 sectors. Directory siblings are chained through `right`.
 */

export type CfbItem =
  | { readonly name: string; readonly data: Uint8Array }
  | { readonly name: string; readonly children: ReadonlyArray<CfbItem> }

const SECTOR = 512
const MINI = 64
const CUTOFF = 4096
const ENDOFCHAIN = 0xFFFFFFFE
const FREESECT = 0xFFFFFFFF
const FATSECT = 0xFFFFFFFD
const DIFSECT = 0xFFFFFFFC
const NOSTREAM = 0xFFFFFFFF

interface Flat {
  readonly name: string
  readonly data: Uint8Array | undefined
  readonly children: Array<number>
  start: number
  right: number
  child: number
}

export interface CfbLayout {
  readonly bytes: Uint8Array
  /** Sector numbers of interest, for tests that patch the file. */
  readonly fatSectors: ReadonlyArray<number>
  readonly difatSectors: ReadonlyArray<number>
  readonly directorySectors: ReadonlyArray<number>
  readonly miniFatSectors: ReadonlyArray<number>
  /** Directory id and first sector of each stream, by path. */
  readonly streams: ReadonlyMap<string, { readonly id: number; readonly start: number; readonly mini: boolean }>
}

export const sectorOffset = (sector: number) => (sector + 1) * SECTOR

export const writeCfb = (items: ReadonlyArray<CfbItem>): CfbLayout => {
  // flatten the tree: entry 0 is the root
  const flat: Array<Flat> = [{ name: "Root Entry", data: undefined, children: [], start: ENDOFCHAIN, right: NOSTREAM, child: NOSTREAM }]
  const paths = new Map<number, string>()
  const add = (list: ReadonlyArray<CfbItem>, parent: number, prefix: string) => {
    for (const item of list) {
      const id = flat.length
      flat.push({ name: item.name, data: "data" in item ? item.data : undefined, children: [], start: ENDOFCHAIN, right: NOSTREAM, child: NOSTREAM })
      flat[parent]!.children.push(id)
      const path = prefix === "" ? item.name : `${prefix}/${item.name}`
      paths.set(id, path)
      if ("children" in item) add(item.children, id, path)
    }
  }
  add(items, 0, "")
  for (const entry of flat) {
    if (entry.children.length > 0) entry.child = entry.children[0]!
    entry.children.forEach((id, k) => {
      flat[id]!.right = entry.children[k + 1] ?? NOSTREAM
    })
  }

  const sectors: Array<Uint8Array> = []
  const fat: Array<number> = []
  /** Appends data as a chain of sectors; returns the first sector. */
  const chain = (data: Uint8Array): number => {
    const count = Math.max(1, Math.ceil(data.length / SECTOR))
    const first = sectors.length
    for (let k = 0; k < count; k++) {
      const s = new Uint8Array(SECTOR)
      s.set(data.subarray(k * SECTOR, (k + 1) * SECTOR))
      sectors.push(s)
      fat.push(k === count - 1 ? ENDOFCHAIN : first + k + 1)
    }
    return first
  }

  // big streams, then the mini stream with the small ones
  const streams = new Map<string, { id: number; start: number; mini: boolean }>()
  const miniParts: Array<Uint8Array> = []
  const miniFat: Array<number> = []
  flat.forEach((entry, id) => {
    if (entry.data === undefined || id === 0) return
    if (entry.data.length >= CUTOFF) {
      entry.start = chain(entry.data)
      streams.set(paths.get(id)!, { id, start: entry.start, mini: false })
    } else if (entry.data.length > 0) {
      const count = Math.ceil(entry.data.length / MINI)
      const first = miniFat.length
      for (let k = 0; k < count; k++) {
        const s = new Uint8Array(MINI)
        s.set(entry.data.subarray(k * MINI, (k + 1) * MINI))
        miniParts.push(s)
        miniFat.push(k === count - 1 ? ENDOFCHAIN : first + k + 1)
      }
      entry.start = first
      streams.set(paths.get(id)!, { id, start: first, mini: true })
    } else {
      streams.set(paths.get(id)!, { id, start: ENDOFCHAIN, mini: true })
    }
  })
  const miniStream = new Uint8Array(miniParts.length * MINI)
  miniParts.forEach((p, k) => miniStream.set(p, k * MINI))
  const root = flat[0]!
  if (miniStream.length > 0) root.start = chain(miniStream)

  // directory
  const dir = new Uint8Array(Math.ceil(flat.length * 128 / SECTOR) * SECTOR)
  const dv = new DataView(dir.buffer)
  flat.forEach((entry, id) => {
    const at = id * 128
    for (let k = 0; k < entry.name.length; k++) dv.setUint16(at + 2 * k, entry.name.charCodeAt(k), true)
    dv.setUint16(at + 64, (entry.name.length + 1) * 2, true)
    dv.setUint8(at + 66, id === 0 ? 5 : entry.data === undefined ? 1 : 2)
    dv.setUint8(at + 67, 1)
    dv.setUint32(at + 68, NOSTREAM, true)
    dv.setUint32(at + 72, entry.right, true)
    dv.setUint32(at + 76, entry.child, true)
    dv.setUint32(at + 116, entry.start, true)
    dv.setUint32(at + 120, id === 0 ? miniStream.length : entry.data?.length ?? 0, true)
  })
  for (let id = flat.length; id < dir.length / 128; id++) {
    dv.setUint32(id * 128 + 68, NOSTREAM, true)
    dv.setUint32(id * 128 + 72, NOSTREAM, true)
    dv.setUint32(id * 128 + 76, NOSTREAM, true)
  }
  const dirStart = chain(dir)
  const directorySectors = Array.from({ length: dir.length / SECTOR }, (_, k) => dirStart + k)

  // miniFAT
  let miniFatStart = ENDOFCHAIN
  let miniFatSectors: Array<number> = []
  if (miniFat.length > 0) {
    const data = new Uint8Array(Math.ceil(miniFat.length * 4 / SECTOR) * SECTOR).fill(0xFF)
    const v = new DataView(data.buffer)
    miniFat.forEach((n, k) => v.setUint32(k * 4, n, true))
    miniFatStart = chain(data)
    miniFatSectors = Array.from({ length: data.length / SECTOR }, (_, k) => miniFatStart + k)
  }

  // FAT and DIFAT sectors: enough to cover every sector including themselves
  const perSector = SECTOR / 4
  let numFat = 1
  let numDifat = 0
  for (;;) {
    numDifat = Math.max(0, Math.ceil((numFat - 109) / (perSector - 1)))
    const needed = Math.ceil((sectors.length + numFat + numDifat) / perSector)
    if (needed <= numFat) break
    numFat = needed
  }
  const fatSectors = Array.from({ length: numFat }, (_, k) => sectors.length + k)
  const difatSectors = Array.from({ length: numDifat }, (_, k) => sectors.length + numFat + k)
  for (let k = 0; k < numFat; k++) fat.push(FATSECT)
  for (let k = 0; k < numDifat; k++) fat.push(DIFSECT)
  const fatData = new Uint8Array(numFat * SECTOR).fill(0xFF)
  const fv = new DataView(fatData.buffer)
  fat.forEach((n, k) => fv.setUint32(k * 4, n, true))
  for (let k = 0; k < numFat; k++) sectors.push(fatData.subarray(k * SECTOR, (k + 1) * SECTOR))
  for (let d = 0; d < numDifat; d++) {
    const s = new Uint8Array(SECTOR).fill(0xFF)
    const sv = new DataView(s.buffer)
    for (let k = 0; k < perSector - 1; k++) {
      const index = 109 + d * (perSector - 1) + k
      if (index < numFat) sv.setUint32(k * 4, fatSectors[index]!, true)
    }
    sv.setUint32(SECTOR - 4, d === numDifat - 1 ? ENDOFCHAIN : difatSectors[d + 1]!, true)
    sectors.push(s)
  }

  // header
  const header = new Uint8Array(SECTOR)
  const hv = new DataView(header.buffer)
  header.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])
  hv.setUint16(24, 0x3E, true)
  hv.setUint16(26, 3, true)
  hv.setUint16(28, 0xFFFE, true)
  hv.setUint16(30, 9, true)
  hv.setUint16(32, 6, true)
  hv.setUint32(44, numFat, true)
  hv.setUint32(48, dirStart, true)
  hv.setUint32(56, CUTOFF, true)
  hv.setUint32(60, miniFatStart, true)
  hv.setUint32(64, miniFatSectors.length, true)
  hv.setUint32(68, numDifat > 0 ? difatSectors[0]! : ENDOFCHAIN, true)
  hv.setUint32(72, numDifat, true)
  for (let k = 0; k < 109; k++) hv.setUint32(76 + 4 * k, k < numFat ? fatSectors[k]! : FREESECT, true)

  const bytes = new Uint8Array(SECTOR * (sectors.length + 1))
  bytes.set(header, 0)
  sectors.forEach((s, k) => bytes.set(s, (k + 1) * SECTOR))
  return { bytes, fatSectors, difatSectors, directorySectors, miniFatSectors, streams }
}

/** Little-endian u32 write into a CFB image. */
export const patch32 = (bytes: Uint8Array, at: number, value: number) => new DataView(bytes.buffer, bytes.byteOffset).setUint32(at, value, true)

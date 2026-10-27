/**
 * Compound File Binary (OLE2 / CFB, MS-CFB) reader: the container of legacy
 * `.xls` workbooks.
 *
 * Opening reads the header, the DIFAT, the FAT, the miniFAT and the
 * directory; streams are read on request. All input is untrusted:
 *
 * - Every sector number is checked against the file before use (a partial
 *   last sector is read as far as it goes).
 * - Every chain (DIFAT, FAT, miniFAT, directory, stream) is cycle-checked
 *   and bounded: by `maxSectorChain`, by the sectors the file has, and for
 *   streams by the sectors their size needs (a walk stops there).
 * - The directory tree is walked once per entry (entries visited twice are
 *   a `MalformedFile`); entry names and sizes are checked.
 *
 * Cost: opening touches each FAT/DIFAT/miniFAT/directory sector once;
 * reading a stream touches its own sectors once. Visited-sector checks use
 * one stamped array per table, so a walk costs its length, not the table
 * size.
 *
 * @internal
 */
import { LimitExceeded, MalformedFile } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"

const SIGNATURE = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
const ENDOFCHAIN = 0xFFFFFFFE
const FREESECT = 0xFFFFFFFF
/** Sector numbers above this are special values (FAT, DIFAT, end, free). */
const MAXREGSECT = 0xFFFFFFFA
const NOSTREAM = 0xFFFFFFFF
const HEADER_DIFAT = 109
const DIRECTORY_ENTRY = 128
const MINI_SECTOR = 64
const MINI_CUTOFF = 4096

export type CfbEntryType = "root" | "storage" | "stream"

export interface CfbEntry {
  /** Index in the directory. */
  readonly id: number
  readonly name: string
  /** Path from the root, `/`-separated (e.g. `Workbook`, `_VBA_PROJECT_CUR/VBA/dir`). */
  readonly path: string
  readonly type: CfbEntryType
  readonly size: number
  readonly start: number
}

export interface Cfb {
  /** Streams and storages under the root, in directory-tree order. */
  readonly entries: ReadonlyArray<CfbEntry>
  /** The entry at `path` (names compared case-insensitively, as CFB does). */
  readonly find: (path: string) => CfbEntry | undefined
  /** A stream's bytes (copied out of its sectors). */
  readonly read: (entry: CfbEntry) => Uint8Array | MalformedFile | LimitExceeded
  /** Sectors (and mini sectors) visited so far: a work counter. */
  readonly sectorsVisited: () => number
}

const malformed = (reason: string, offset?: number) =>
  new MalformedFile({ format: "cfb", reason, ...(offset === undefined ? {} : { offset }) })

/**
 * Walks a sector chain with a table (`FAT` or miniFAT).
 *
 * Progress: each step moves to a sector not visited in this walk (checked
 * with a stamp per walk), so a walk ends after at most `table.length`
 * steps; it also ends at `maxSteps` (the sectors the caller needs, or
 * `maxSectorChain`).
 */
class Chains {
  private readonly stamps: Uint32Array
  private stamp = 0
  visited = 0
  private readonly table: Uint32Array
  private readonly limitSectors: number
  private readonly maxChain: number
  private readonly what: string

  constructor(table: Uint32Array, limitSectors: number, maxChain: number, what: string) {
    this.table = table
    this.stamps = new Uint32Array(table.length)
    this.limitSectors = limitSectors
    this.maxChain = maxChain
    this.what = what
  }

  /** Up to `count` sectors of the chain from `start` (all of them when `count` is undefined). */
  walk(start: number, count: number | undefined): Array<number> | MalformedFile | LimitExceeded {
    const out: Array<number> = []
    if (count === 0) return out
    this.stamp++
    let sector = start
    // progress: every step adds one new (unvisited) sector, see the class doc
    while (sector !== ENDOFCHAIN) {
      if (sector > MAXREGSECT || sector >= this.limitSectors || sector >= this.table.length) {
        return malformed(`${this.what} chain points at sector ${sector}`)
      }
      if (this.stamps[sector] === this.stamp) return malformed(`${this.what} chain has a cycle at sector ${sector}`)
      if (out.length >= this.maxChain) {
        return new LimitExceeded({ limit: "maxSectorChain", max: this.maxChain, actual: out.length + 1 })
      }
      this.stamps[sector] = this.stamp
      out.push(sector)
      this.visited++
      if (count !== undefined && out.length === count) return out
      sector = this.table[sector]!
    }
    return out
  }
}

const utf16 = new TextDecoder("utf-16le")

/** Opens a compound file. */
export const open = (bytes: Uint8Array, limits: Limits): Cfb | MalformedFile | LimitExceeded => {
  if (bytes.length < 512) return malformed("file shorter than a CFB header")
  // progress: k up to 8
  for (let k = 0; k < 8; k++) if (bytes[k] !== SIGNATURE[k]) return malformed("not a CFB file (bad signature)", 0)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const u16 = (at: number) => view.getUint16(at, true)
  const u32 = (at: number) => view.getUint32(at, true)
  if (u16(28) !== 0xFFFE) return malformed("bad byte order mark", 28)
  const major = u16(26)
  const shift = u16(30)
  if (!((major === 3 && shift === 9) || (major === 4 && shift === 12))) {
    return malformed(`unsupported version ${major} with sector size 2^${shift}`, 26)
  }
  if (u16(32) !== 6) return malformed("mini sector size must be 64", 32)
  if (u32(56) !== MINI_CUTOFF) return malformed("mini stream cutoff must be 4096", 56)
  const sectorSize = 1 << shift
  if (bytes.length < sectorSize) return malformed("file shorter than its header sector")
  /** Sectors the file holds (a partial last one counts). */
  const fileSectors = Math.ceil((bytes.length - sectorSize) / sectorSize)
  const perSector = sectorSize / 4
  const maxChain = limits.maxSectorChain
  let visited = 0

  const sectorAt = (sector: number): Uint8Array => {
    const at = (sector + 1) * sectorSize
    return bytes.subarray(at, Math.min(bytes.length, at + sectorSize))
  }
  const sectorU32 = (sector: number, index: number): number => {
    const at = (sector + 1) * sectorSize + index * 4
    return at + 4 <= bytes.length ? view.getUint32(at, true) : FREESECT
  }

  // FAT sector numbers: 109 in the header, the rest in the DIFAT chain
  const numFat = u32(44)
  if (numFat > fileSectors) return malformed(`header lists ${numFat} FAT sectors, the file has ${fileSectors}`, 44)
  const fatSectors: Array<number> = []
  // progress: k up to 109
  for (let k = 0; k < HEADER_DIFAT && fatSectors.length < numFat; k++) fatSectors.push(u32(76 + 4 * k))
  let difat = u32(68)
  const numDifat = u32(72)
  const difatSeen = new Set<number>()
  // progress: each DIFAT sector is visited once (difatSeen), at most numDifat and fileSectors of them
  while (fatSectors.length < numFat && difat !== ENDOFCHAIN && difat !== FREESECT) {
    if (difat > MAXREGSECT || difat >= fileSectors) return malformed(`DIFAT chain points at sector ${difat}`)
    if (difatSeen.has(difat)) return malformed(`DIFAT chain has a cycle at sector ${difat}`)
    if (difatSeen.size >= numDifat) return malformed("DIFAT chain is longer than the header says")
    if (difatSeen.size >= maxChain) return new LimitExceeded({ limit: "maxSectorChain", max: maxChain, actual: difatSeen.size + 1 })
    difatSeen.add(difat)
    visited++
    // progress: k up to perSector - 1
    for (let k = 0; k < perSector - 1 && fatSectors.length < numFat; k++) fatSectors.push(sectorU32(difat, k))
    difat = sectorU32(difat, perSector - 1)
  }
  if (fatSectors.length < numFat) return malformed(`found ${fatSectors.length} of ${numFat} FAT sectors`)

  // the FAT itself
  const fat = new Uint32Array(numFat * perSector)
  // progress: one FAT sector per iteration
  for (let k = 0; k < numFat; k++) {
    const sector = fatSectors[k]!
    if (sector > MAXREGSECT || sector >= fileSectors) return malformed(`FAT sector ${sector} is outside the file`)
    const data = sectorAt(sector)
    visited++
    const words = Math.floor(data.length / 4)
    // progress: i up to words
    for (let i = 0; i < words; i++) fat[k * perSector + i] = view.getUint32(data.byteOffset - bytes.byteOffset + i * 4, true)
    // a partial last sector leaves the rest free
    for (let i = words; i < perSector; i++) fat[k * perSector + i] = FREESECT
  }
  const chains = new Chains(fat, fileSectors, maxChain, "FAT")

  /** Copies `size` bytes of a chain of `unit`-byte sectors. */
  const gather = (
    sectors: ReadonlyArray<number>,
    size: number,
    unit: number,
    slice: (sector: number) => Uint8Array
  ): Uint8Array | MalformedFile => {
    const out = new Uint8Array(size)
    let at = 0
    // progress: one sector per iteration
    for (const sector of sectors) {
      if (at >= size) break
      const data = slice(sector)
      const take = Math.min(unit, size - at)
      if (data.length < take) return malformed("stream runs past the end of the file")
      out.set(data.subarray(0, take), at)
      at += take
    }
    if (at < size) return malformed("stream chain is shorter than its size")
    return out
  }

  // the directory
  const dirChain = chains.walk(u32(48), undefined)
  if (!Array.isArray(dirChain)) return dirChain
  const dirBytes = gather(dirChain, dirChain.length * sectorSize, sectorSize, sectorAt)
  if (dirBytes instanceof MalformedFile) return dirBytes
  const dirView = new DataView(dirBytes.buffer)
  const count = Math.floor(dirBytes.length / DIRECTORY_ENTRY)
  if (count === 0) return malformed("empty directory")

  interface Raw {
    readonly name: string
    readonly type: number
    readonly left: number
    readonly right: number
    readonly child: number
    readonly start: number
    readonly size: number
  }
  const raw = (id: number): Raw | MalformedFile => {
    const at = id * DIRECTORY_ENTRY
    const nameLength = dirView.getUint16(at + 64, true)
    if (nameLength > 64 || nameLength % 2 !== 0) return malformed(`directory entry ${id} has a bad name length`)
    const name = utf16.decode(dirBytes.subarray(at, at + Math.max(0, nameLength - 2)))
    const low = dirView.getUint32(at + 120, true)
    const high = dirView.getUint32(at + 124, true)
    // version 3 files may leave garbage in the high word
    const size = major === 3 ? low : high >= 0x200000 ? Infinity : high * 0x100000000 + low
    return {
      name,
      type: dirView.getUint8(at + 66),
      left: dirView.getUint32(at + 68, true),
      right: dirView.getUint32(at + 72, true),
      child: dirView.getUint32(at + 76, true),
      start: dirView.getUint32(at + 116, true),
      size
    }
  }
  const root = raw(0)
  if (root instanceof MalformedFile) return root
  if (root.type !== 5) return malformed("the first directory entry is not the root")

  // walk the tree: siblings (left/right) and children; every entry once
  const entries: Array<CfbEntry> = []
  const seen = new Uint8Array(count)
  seen[0] = 1
  const stack: Array<readonly [number, string]> = root.child === NOSTREAM ? [] : [[root.child, ""]]
  // progress: each entry is pushed at most once more than it is visited,
  // and visited at most once (seen), so at most `count` iterations do work
  while (stack.length > 0) {
    const [id, prefix] = stack.pop()!
    if (id >= count) return malformed(`directory points at entry ${id} of ${count}`)
    if (seen[id] === 1) return malformed(`directory entry ${id} is reached twice`)
    seen[id] = 1
    const entry = raw(id)
    if (entry instanceof MalformedFile) return entry
    const type: CfbEntryType | undefined = entry.type === 1 ? "storage" : entry.type === 2 ? "stream" : undefined
    if (type === undefined) return malformed(`directory entry ${id} has type ${entry.type}`)
    const path = prefix === "" ? entry.name : `${prefix}/${entry.name}`
    entries.push({ id, name: entry.name, path, type, size: type === "stream" ? entry.size : 0, start: entry.start })
    if (entry.right !== NOSTREAM) stack.push([entry.right, prefix])
    if (entry.left !== NOSTREAM) stack.push([entry.left, prefix])
    if (type === "storage" && entry.child !== NOSTREAM) stack.push([entry.child, path])
  }
  entries.sort((a, b) => a.id - b.id)

  // the mini stream (in the root entry) and the miniFAT, read on first use
  let mini: { readonly data: Uint8Array; readonly chains: Chains } | MalformedFile | LimitExceeded | undefined
  const loadMini = () => {
    if (mini !== undefined) return mini
    const numMiniFat = u32(64)
    const miniFatChain = chains.walk(u32(60), numMiniFat)
    if (!Array.isArray(miniFatChain)) return (mini = miniFatChain)
    if (miniFatChain.length < numMiniFat) return (mini = malformed("miniFAT chain is shorter than the header says"))
    const table = new Uint32Array(miniFatChain.length * perSector)
    // progress: one sector per iteration, i up to perSector
    miniFatChain.forEach((sector, k) => {
      for (let i = 0; i < perSector; i++) table[k * perSector + i] = sectorU32(sector, i)
    })
    if (!Number.isFinite(root.size) || root.size > fileSectors * sectorSize) {
      return (mini = malformed("mini stream is larger than the file"))
    }
    const rootChain = chains.walk(root.start, Math.ceil(root.size / sectorSize))
    if (!Array.isArray(rootChain)) return (mini = rootChain)
    const data = gather(rootChain, root.size, sectorSize, sectorAt)
    if (data instanceof MalformedFile) return (mini = data)
    const miniSectors = Math.floor(data.length / MINI_SECTOR)
    return (mini = { data, chains: new Chains(table, miniSectors, maxChain, "miniFAT") })
  }

  const byPath = new Map(entries.map((e) => [e.path.toLowerCase(), e]))
  return {
    entries,
    find: (path) => byPath.get(path.toLowerCase()),
    read: (entry) => {
      if (entry.type !== "stream") return malformed(`${entry.path} is not a stream`)
      const size = entry.size
      if (!Number.isFinite(size) || size > fileSectors * sectorSize) {
        return malformed(`stream ${entry.path} is larger than the file`)
      }
      if (size < MINI_CUTOFF) {
        if (size === 0) return new Uint8Array(0)
        const m = loadMini()
        if (!("data" in m)) return m
        const sectors = m.chains.walk(entry.start, Math.ceil(size / MINI_SECTOR))
        if (!Array.isArray(sectors)) return sectors
        return gather(sectors, size, MINI_SECTOR, (s) => m.data.subarray(s * MINI_SECTOR, s * MINI_SECTOR + MINI_SECTOR))
      }
      const sectors = chains.walk(entry.start, Math.ceil(size / sectorSize))
      if (!Array.isArray(sectors)) return sectors
      return gather(sectors, size, sectorSize, sectorAt)
    },
    sectorsVisited: () => {
      const m = mini !== undefined && "data" in mini ? mini.chains.visited : 0
      return visited + chains.visited + m
    }
  }
}

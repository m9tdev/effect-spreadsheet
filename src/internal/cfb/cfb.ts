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
import * as Result from "effect/Result"
import { LimitExceeded, MalformedFile } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { type Meter, noMeter } from "../meter.ts"

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
  /** Streams and storages under the root, by directory id. */
  readonly entries: ReadonlyArray<CfbEntry>
  /**
   * The entry at `path` (names compared case-insensitively, as CFB does).
   * Names can't contain `/` (opening rejects them); if two entries have the
   * same path, the one with the lower directory id wins.
   */
  readonly find: (path: string) => CfbEntry | undefined
  /** A stream's bytes (copied out of its sectors). */
  readonly read: (entry: CfbEntry) => Result.Result<Uint8Array, CfbError>
}

/** What opening a compound file or reading one of its streams fails with. */
export type CfbError = MalformedFile | LimitExceeded

const malformed = (reason: string, offset?: number) =>
  new MalformedFile({ format: "cfb", reason, ...(offset === undefined ? {} : { offset }) })

/** A failed result with a `MalformedFile`. */
const fail = (reason: string, offset?: number) => Result.fail(malformed(reason, offset))

/**
 * Walks a sector chain with a table (`FAT` or miniFAT).
 *
 * Progress: each step moves to a sector not visited in this walk (checked
 * with a stamp per walk), so a walk ends after at most `table.length`
 * steps; it also ends at `maxSteps` (the sectors the caller needs, or
 * `maxSectorChain`). The sectors each walk visits go to the work meter.
 */
class Chains {
  /**
   * Who owns each sector: 0 none, `id + 1` for directory entry `id`, a
   * negative value for the container's own structures. A sector may be
   * claimed once: two streams sharing sectors would let reading every
   * stream copy the same data again and again.
   */
  private readonly owners: Int32Array
  private readonly stamps: Uint32Array
  private stamp = 0
  private readonly table: Uint32Array
  private readonly limitSectors: number
  private readonly maxChain: number
  private readonly what: string
  private readonly meter: Meter

  constructor(table: Uint32Array, limitSectors: number, maxChain: number, what: string, meter: Meter) {
    this.table = table
    this.stamps = new Uint32Array(table.length)
    this.owners = new Int32Array(table.length)
    this.limitSectors = limitSectors
    this.maxChain = maxChain
    this.what = what
    this.meter = meter
  }

  /** Marks a sector as the container's own (FAT, DIFAT). */
  reserve(sector: number, owner: number): void {
    if (sector < this.owners.length) this.owners[sector] = owner
  }

  /**
   * Up to `count` sectors of the chain from `start` (all of them when
   * `count` is undefined), claimed for `owner`: a sector another owner has
   * claimed is a `MalformedFile` (the same owner may walk it again).
   */
  walk(start: number, count: number | undefined, owner: number): Result.Result<Array<number>, CfbError> {
    const out: Array<number> = []
    if (count === 0) return Result.succeed(out)
    const result = this.follow(start, count, owner, out)
    this.meter.add("sectors", out.length)
    return result
  }

  /** `walk`'s loop: collects the chain's sectors in `out`. */
  private follow(start: number, count: number | undefined, owner: number, out: Array<number>): Result.Result<Array<number>, CfbError> {
    this.stamp++
    let sector = start
    // progress: every step adds one new (unvisited) sector, see the class doc
    while (sector !== ENDOFCHAIN) {
      if (sector > MAXREGSECT || sector >= this.limitSectors || sector >= this.table.length) {
        return fail(`${this.what} chain points at sector ${sector}`)
      }
      if (this.stamps[sector] === this.stamp) return fail(`${this.what} chain has a cycle at sector ${sector}`)
      if (out.length >= this.maxChain) {
        return Result.fail(new LimitExceeded({ limit: "maxSectorChain", max: this.maxChain, actual: out.length + 1 }))
      }
      const holder = this.owners[sector]!
      if (holder !== 0 && holder !== owner) return fail(`sector ${sector} of the ${this.what} belongs to two streams`)
      this.owners[sector] = owner
      this.stamps[sector] = this.stamp
      out.push(sector)
      if (count !== undefined && out.length === count) return Result.succeed(out)
      sector = this.table[sector]!
    }
    return Result.succeed(out)
  }
}

const utf16 = new TextDecoder("utf-16le")

/** Whether each sector follows the one before it. */
const consecutive = (sectors: ReadonlyArray<number>): boolean => {
  // progress: k up to sectors.length
  for (let k = 1; k < sectors.length; k++) if (sectors[k] !== sectors[k - 1]! + 1) return false
  return true
}

/** Opens a compound file. The sectors it visits, opening and reading streams, go to `meter`. */
export const open = (bytes: Uint8Array, limits: Limits, meter: Meter = noMeter): Result.Result<Cfb, CfbError> => {
  if (bytes.length < 512) return fail("file shorter than a CFB header")
  // progress: k up to 8
  for (let k = 0; k < 8; k++) if (bytes[k] !== SIGNATURE[k]) return fail("not a CFB file (bad signature)", 0)
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const u16 = (at: number) => view.getUint16(at, true)
  const u32 = (at: number) => view.getUint32(at, true)
  if (u16(28) !== 0xFFFE) return fail("bad byte order mark", 28)
  const major = u16(26)
  const shift = u16(30)
  if (!((major === 3 && shift === 9) || (major === 4 && shift === 12))) {
    return fail(`unsupported version ${major} with sector size 2^${shift}`, 26)
  }
  if (u16(32) !== 6) return fail("mini sector size must be 64", 32)
  if (u32(56) !== MINI_CUTOFF) return fail("mini stream cutoff must be 4096", 56)
  const sectorSize = 1 << shift
  if (bytes.length < sectorSize) return fail("file shorter than its header sector")
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

  // FAT sector numbers: 109 in the header, the rest in the DIFAT chain.
  // Some writers (pyExcelerator) declare more FAT sectors than the file has
  // (109 in an 8-sector file) and list only the real ones, the rest free.
  // The declared count only bounds how many numbers are collected (at most
  // one per sector of the file); the FAT is the in-use entries listed. A chain
  // that reaches past the loaded FAT fails when it is walked.
  const declaredFat = u32(44)
  const numFatWanted = Math.min(declaredFat, fileSectors)
  const fatSectors: Array<number> = []
  let headerEnded = false
  // progress: k up to 109
  for (let k = 0; k < HEADER_DIFAT && fatSectors.length < numFatWanted; k++) {
    const sector = u32(76 + 4 * k)
    if (sector === FREESECT) {
      headerEnded = true
      break
    }
    fatSectors.push(sector)
  }
  let difat = u32(68)
  const numDifat = u32(72)
  const difatSeen = new Set<number>()
  // progress: each DIFAT sector is visited once (difatSeen), at most numDifat and fileSectors of them
  while (!headerEnded && fatSectors.length < numFatWanted && difat !== ENDOFCHAIN && difat !== FREESECT) {
    if (difat > MAXREGSECT || difat >= fileSectors) return fail(`DIFAT chain points at sector ${difat}`)
    if (difatSeen.has(difat)) return fail(`DIFAT chain has a cycle at sector ${difat}`)
    if (difatSeen.size >= numDifat) return fail("DIFAT chain is longer than the header says")
    if (difatSeen.size >= maxChain) return Result.fail(new LimitExceeded({ limit: "maxSectorChain", max: maxChain, actual: difatSeen.size + 1 }))
    difatSeen.add(difat)
    visited++
    // progress: k up to perSector - 1
    for (let k = 0; k < perSector - 1 && fatSectors.length < numFatWanted; k++) {
      const sector = sectorU32(difat, k)
      if (sector === FREESECT) {
        headerEnded = true // the listed FAT sectors end here
        break
      }
      fatSectors.push(sector)
    }
    if (headerEnded) break
    difat = sectorU32(difat, perSector - 1)
  }
  if (fatSectors.length === 0) return fail("no FAT sectors")
  // the FAT itself. FAT sector k describes sectors k * perSector onwards, so
  // only the first ceil(fileSectors / perSector) can describe sectors the
  // file has; walks reject sectors past the file before using the FAT, so the
  // rest are never consulted and are not read. (pyExcelerator lists FAT
  // sectors past the end of a small file; those must not fail it, while a
  // FAT sector a chain could need and the file lacks still does.)
  const numFat = Math.min(fatSectors.length, Math.ceil(fileSectors / perSector))
  const fat = new Uint32Array(numFat * perSector)
  // progress: one FAT sector per iteration
  for (let k = 0; k < numFat; k++) {
    const sector = fatSectors[k]!
    if (sector > MAXREGSECT || sector >= fileSectors) return fail(`FAT sector ${sector} is outside the file`)
    const data = sectorAt(sector)
    visited++
    const words = Math.floor(data.length / 4)
    // progress: i up to words
    for (let i = 0; i < words; i++) fat[k * perSector + i] = view.getUint32(data.byteOffset - bytes.byteOffset + i * 4, true)
    // a partial last sector leaves the rest free
    for (let i = words; i < perSector; i++) fat[k * perSector + i] = FREESECT
  }
  meter.add("sectors", visited)
  const chains = new Chains(fat, fileSectors, maxChain, "FAT", meter)
  const OWN_FAT = -1
  const OWN_DIRECTORY = -2
  const OWN_MINIFAT = -3
  // progress: one sector per iteration
  for (const sector of fatSectors) chains.reserve(sector, OWN_FAT)
  for (const sector of difatSeen) chains.reserve(sector, OWN_FAT)

  /** Copies `size` bytes of a chain of `unit`-byte sectors. */
  const gather = (
    sectors: ReadonlyArray<number>,
    size: number,
    unit: number,
    slice: (sector: number) => Uint8Array
  ): Result.Result<Uint8Array, MalformedFile> => {
    const out = new Uint8Array(size)
    let at = 0
    // progress: one sector per iteration
    for (const sector of sectors) {
      if (at >= size) break
      const data = slice(sector)
      const take = Math.min(unit, size - at)
      if (data.length < take) return fail("stream runs past the end of the file")
      out.set(data.subarray(0, take), at)
      at += take
    }
    if (at < size) return fail("stream chain is shorter than its size")
    return Result.succeed(out)
  }

  // the directory
  const dirChain = chains.walk(u32(48), undefined, OWN_DIRECTORY)
  if (Result.isFailure(dirChain)) return Result.fail(dirChain.failure)
  const dirGathered = gather(dirChain.success, dirChain.success.length * sectorSize, sectorSize, sectorAt)
  if (Result.isFailure(dirGathered)) return Result.fail(dirGathered.failure)
  const dirBytes = dirGathered.success
  const dirView = new DataView(dirBytes.buffer)
  const count = Math.floor(dirBytes.length / DIRECTORY_ENTRY)
  if (count === 0) return fail("empty directory")

  interface Raw {
    readonly name: string
    readonly type: number
    readonly left: number
    readonly right: number
    readonly child: number
    readonly start: number
    readonly size: number
  }
  const raw = (id: number): Result.Result<Raw, MalformedFile> => {
    const at = id * DIRECTORY_ENTRY
    const nameLength = dirView.getUint16(at + 64, true)
    if (nameLength > 64 || nameLength % 2 !== 0) return fail(`directory entry ${id} has a bad name length`)
    const name = utf16.decode(dirBytes.subarray(at, at + Math.max(0, nameLength - 2)))
    const low = dirView.getUint32(at + 120, true)
    const high = dirView.getUint32(at + 124, true)
    // version 3 files may leave garbage in the high word
    const size = major === 3 ? low : high >= 0x200000 ? Infinity : high * 0x100000000 + low
    return Result.succeed({
      name,
      type: dirView.getUint8(at + 66),
      left: dirView.getUint32(at + 68, true),
      right: dirView.getUint32(at + 72, true),
      child: dirView.getUint32(at + 76, true),
      start: dirView.getUint32(at + 116, true),
      size
    })
  }
  const rawRoot = raw(0)
  if (Result.isFailure(rawRoot)) return Result.fail(rawRoot.failure)
  const root = rawRoot.success
  if (root.type !== 5) return fail("the first directory entry is not the root")

  // walk the tree: siblings (left/right) and children; every entry once
  const entries: Array<CfbEntry> = []
  const seen = new Uint8Array(count)
  seen[0] = 1
  // [id, parent path, depth]; paths are built per entry, so the depth is
  // bounded (maxStorageDepth) to keep their total length in check
  const stack: Array<readonly [number, string, number]> = root.child === NOSTREAM ? [] : [[root.child, "", 1]]
  // progress: each entry is pushed at most once more than it is visited,
  // and visited at most once (seen), so at most `count` iterations do work
  while (stack.length > 0) {
    const [id, prefix, depth] = stack.pop()!
    if (id >= count) return fail(`directory points at entry ${id} of ${count}`)
    if (seen[id] === 1) return fail(`directory entry ${id} is reached twice`)
    seen[id] = 1
    const rawEntry = raw(id)
    if (Result.isFailure(rawEntry)) return Result.fail(rawEntry.failure)
    const entry = rawEntry.success
    const type: CfbEntryType | undefined = entry.type === 1 ? "storage" : entry.type === 2 ? "stream" : undefined
    if (type === undefined) return fail(`directory entry ${id} has type ${entry.type}`)
    // MS-CFB forbids "/" (and \ : !) in names; "/" would break paths
    if (entry.name.includes("/")) return fail(`directory entry ${id} has "/" in its name`)
    const path = prefix === "" ? entry.name : `${prefix}/${entry.name}`
    entries.push({ id, name: entry.name, path, type, size: type === "stream" ? entry.size : 0, start: entry.start })
    if (entry.right !== NOSTREAM) stack.push([entry.right, prefix, depth])
    if (entry.left !== NOSTREAM) stack.push([entry.left, prefix, depth])
    if (type === "storage" && entry.child !== NOSTREAM) {
      if (depth >= limits.maxStorageDepth) {
        return Result.fail(new LimitExceeded({ limit: "maxStorageDepth", max: limits.maxStorageDepth, actual: depth + 1 }))
      }
      stack.push([entry.child, path, depth + 1])
    }
  }
  entries.sort((a, b) => a.id - b.id)

  // the mini stream (in the root entry) and the miniFAT, read on first use
  interface Mini {
    readonly data: Uint8Array
    readonly chains: Chains
  }
  let mini: Result.Result<Mini, CfbError> | undefined
  const loadMini = (): Result.Result<Mini, CfbError> => {
    if (mini !== undefined) return mini
    const numMiniFat = u32(64)
    const miniFatChain = chains.walk(u32(60), numMiniFat, OWN_MINIFAT)
    if (Result.isFailure(miniFatChain)) return (mini = Result.fail(miniFatChain.failure))
    const miniFatSectors = miniFatChain.success
    if (miniFatSectors.length < numMiniFat) return (mini = fail("miniFAT chain is shorter than the header says"))
    const table = new Uint32Array(miniFatSectors.length * perSector)
    // progress: one sector per iteration, i up to perSector
    miniFatSectors.forEach((sector, k) => {
      for (let i = 0; i < perSector; i++) table[k * perSector + i] = sectorU32(sector, i)
    })
    if (!Number.isFinite(root.size) || root.size > fileSectors * sectorSize) {
      return (mini = fail("mini stream is larger than the file"))
    }
    const rootChain = chains.walk(root.start, Math.ceil(root.size / sectorSize), 1)
    if (Result.isFailure(rootChain)) return (mini = Result.fail(rootChain.failure))
    const data = gather(rootChain.success, root.size, sectorSize, sectorAt)
    if (Result.isFailure(data)) return (mini = Result.fail(data.failure))
    // a last, partial mini sector counts (its stream reads only what it has)
    const miniSectors = Math.ceil(data.success.length / MINI_SECTOR)
    return (mini = Result.succeed({ data: data.success, chains: new Chains(table, miniSectors, maxChain, "miniFAT", meter) }))
  }

  const byPath = new Map<string, CfbEntry>()
  // progress: one entry per iteration; entries are sorted by id, so the first wins
  for (const entry of entries) {
    const key = entry.path.toLowerCase()
    if (!byPath.has(key)) byPath.set(key, entry)
  }
  return Result.succeed({
    entries,
    find: (path) => byPath.get(path.toLowerCase()),
    read: (entry) => {
      if (entry.type !== "stream") return fail(`${entry.path} is not a stream`)
      const size = entry.size
      if (!Number.isFinite(size) || size > fileSectors * sectorSize) {
        return fail(`stream ${entry.path} is larger than the file`)
      }
      if (size < MINI_CUTOFF) {
        if (size === 0) return Result.succeed(new Uint8Array(0))
        const m = loadMini()
        if (Result.isFailure(m)) return Result.fail(m.failure)
        const { data, chains: miniChains } = m.success
        const sectors = miniChains.walk(entry.start, Math.ceil(size / MINI_SECTOR), entry.id + 1)
        if (Result.isFailure(sectors)) return Result.fail(sectors.failure)
        return gather(sectors.success, size, MINI_SECTOR, (s) => data.subarray(s * MINI_SECTOR, s * MINI_SECTOR + MINI_SECTOR))
      }
      const walked = chains.walk(entry.start, Math.ceil(size / sectorSize), entry.id + 1)
      if (Result.isFailure(walked)) return Result.fail(walked.failure)
      const sectors = walked.success
      // a stream in consecutive sectors (as writers lay out big streams) is a
      // view of the file, not a copy: no pause to copy hundreds of MiB
      const at = sectors.length > 0 ? (sectors[0]! + 1) * sectorSize : 0
      if (sectors.length * sectorSize >= size && at + size <= bytes.length && consecutive(sectors)) {
        return Result.succeed(bytes.subarray(at, at + size))
      }
      return gather(sectors, size, sectorSize, sectorAt)
    }
  })
}

/**
 * Compound File Binary (OLE2 / CFB, MS-CFB) reader: the container of legacy
 * `.xls` workbooks.
 *
 * Opening reads the header, the DIFAT, the FAT, the miniFAT and the
 * directory; streams are read on request. All input is untrusted:
 *
 * - Every sector number is checked against the file before use (a partial
 *   last sector is read as far as it goes).
 * - Every chain (DIFAT, FAT, miniFAT, directory, stream) is walked by one
 *   walker (`walk`): cycle-checked and bounded by `maxSectorChain`, by the
 *   sectors the file has, and for streams by the sectors their size needs
 *   (a walk stops there).
 * - Each sector has at most one owner (the container's own structures, or
 *   one stream): streams that share sectors are a `MalformedFile`.
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

/**
 * A stream's size: its byte count, or `Invalid` when the directory gives a
 * size larger than the file. An entry with an invalid size opens; reading
 * its stream fails.
 */
export type CfbSize =
  | { readonly _tag: "Valid"; readonly bytes: number }
  | { readonly _tag: "Invalid" }

export interface CfbEntry {
  /** Index in the directory. */
  readonly id: number
  readonly name: string
  /** Path from the root, `/`-separated (e.g. `Workbook`, `_VBA_PROJECT_CUR/VBA/dir`). */
  readonly path: string
  readonly type: CfbEntryType
  /** The stream's size (0 for a storage). */
  readonly size: CfbSize
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
  /**
   * A stream's bytes: a view of the file when the stream is in consecutive
   * sectors of it, else a copy gathered from its sectors (mini streams are
   * always copies). Either way they come from the caller's buffer when
   * read: the mini stream, loaded on the first read of a mini stream, is a
   * view of it too when its sectors are consecutive. So changing the
   * buffer between reads changes what later reads give, for mini streams
   * as for big ones.
   */
  readonly read: (entry: CfbEntry) => Result.Result<Uint8Array, CfbError>
}

/** What opening a compound file or reading one of its streams fails with. */
export type CfbError = MalformedFile | LimitExceeded

const malformed = (reason: string, offset?: number) =>
  new MalformedFile({ format: "cfb", reason, ...(offset === undefined ? {} : { offset }) })

/** A failed result with a `MalformedFile`. */
const fail = (reason: string, offset?: number) => Result.fail(malformed(reason, offset))

/**
 * Who holds a sector: the container's own structures (the FAT and DIFAT
 * sectors, the directory, the miniFAT), or the stream of a directory entry
 * (the root entry's stream is the mini stream).
 */
type Owner =
  | { readonly _tag: "Fat" }
  | { readonly _tag: "Directory" }
  | { readonly _tag: "MiniFat" }
  | { readonly _tag: "Stream"; readonly id: number }

const FAT_OWNER: Owner = { _tag: "Fat" }
const DIRECTORY_OWNER: Owner = { _tag: "Directory" }
const MINIFAT_OWNER: Owner = { _tag: "MiniFat" }
const streamOwner = (id: number): Owner => ({ _tag: "Stream", id })

/** An owner as one number, for the owner table: distinct per owner, never 0. */
const ownerKey = (owner: Owner): number => {
  switch (owner._tag) {
    case "Fat":
      return -1
    case "Directory":
      return -2
    case "MiniFat":
      return -3
    case "Stream":
      return owner.id + 1
  }
}

/**
 * The sectors one table numbers (the file's, or the mini stream's): which
 * ones the current walk has visited (a stamp per walk), and who owns each.
 */
class Sectors {
  private readonly stamps: Uint32Array
  private stamp = 0
  /**
   * Each sector's owner, by `ownerKey` (0: none). A sector may be claimed by
   * one owner: two streams sharing sectors would let reading every stream
   * copy the same data again and again.
   */
  private readonly owners: Int32Array

  constructor(length: number) {
    this.stamps = new Uint32Array(length)
    this.owners = new Int32Array(length)
  }

  /** Starts a walk: no sector is visited yet. */
  begin(): void {
    this.stamp++
  }

  /** Marks `sector` visited in this walk; false if it already was. */
  visit(sector: number): boolean {
    if (this.stamps[sector] === this.stamp) return false
    this.stamps[sector] = this.stamp
    return true
  }

  /** Marks a sector as `owner`'s without a check (the FAT and DIFAT sectors, listed by the header). */
  reserve(sector: number, owner: Owner): void {
    if (sector < this.owners.length) this.owners[sector] = ownerKey(owner)
  }

  /** Claims `sector` for the owner with `key`; false if another owner holds it (the same owner may claim it again). */
  claim(sector: number, key: number): boolean {
    const holder = this.owners[sector]!
    if (holder !== 0 && holder !== key) return false
    this.owners[sector] = key
    return true
  }
}

/**
 * How a chain goes on: a table (the FAT or the miniFAT) that gives the
 * sector after each sector, or a function (the DIFAT's: the last word of a
 * DIFAT sector). `ENDOFCHAIN` ends it.
 */
type Next = Uint32Array | ((sector: number) => number)

interface WalkOptions {
  /** The chain's table, for messages: `FAT`, `miniFAT` or `DIFAT`. */
  readonly what: string
  /** Sector numbers from this one on are outside (the file, the mini stream, the table). */
  readonly limitSectors: number
  /** `maxSectorChain`. */
  readonly max: number
  /** Stop after this many sectors (the ones a stream's size needs); all of them when undefined. */
  readonly count?: number | undefined
  /** The chain length the header declares: a longer chain is a `MalformedFile`. */
  readonly declared?: number | undefined
  /** Who the walked sectors are claimed for (none: not claimed). */
  readonly owner?: Owner | undefined
  readonly sectors: Sectors
  /** The sectors walked go to it. */
  readonly meter: Meter
}

/**
 * The sectors of the chain from `start`, each after the one before by `next`.
 *
 * Progress: each step moves to a sector not visited in this walk (checked
 * with a stamp per walk), so a walk ends after at most `limitSectors` steps;
 * it also ends at `max` (and at `count`, `declared`).
 */
const walk = (start: number, next: Next, o: WalkOptions): Result.Result<Array<number>, CfbError> => {
  const out: Array<number> = []
  if (o.count === 0) return Result.succeed(out)
  const result = follow(start, next, o, out)
  o.meter.add("sectors", out.length)
  return result
}

/** `walk`'s loop: collects the chain's sectors in `out`. */
const follow = (start: number, next: Next, o: WalkOptions, out: Array<number>): Result.Result<Array<number>, CfbError> => {
  // the options as locals: the loop runs once per sector
  const { what, sectors, limitSectors, max } = o
  const declared = o.declared ?? Infinity
  const count = o.count ?? Infinity
  const owner = o.owner === undefined ? 0 : ownerKey(o.owner)
  sectors.begin()
  let sector = start
  // progress: every step adds one new (unvisited) sector, see `walk`
  while (sector !== ENDOFCHAIN) {
    if (sector > MAXREGSECT || sector >= limitSectors) return fail(`${what} chain points at sector ${sector}`)
    if (!sectors.visit(sector)) return fail(`${what} chain has a cycle at sector ${sector}`)
    if (out.length >= declared) return fail(`${what} chain is longer than the header says`)
    if (out.length >= max) return Result.fail(new LimitExceeded({ limit: "maxSectorChain", max, actual: out.length + 1 }))
    if (owner !== 0 && !sectors.claim(sector, owner)) return fail(`sector ${sector} of the ${what} belongs to two streams`)
    out.push(sector)
    if (out.length === count) return Result.succeed(out)
    // a table is read here, not through a function call: this runs once per sector
    sector = typeof next === "function" ? next(sector) : next[sector]!
  }
  return Result.succeed(out)
}

/**
 * Where a stream's sectors are and how its chain goes: the file's sectors
 * with the FAT, or the mini stream's mini sectors with the miniFAT.
 */
interface Medium {
  readonly what: "FAT" | "miniFAT"
  /** The bytes sector 0 starts in, at `first`. */
  readonly data: Uint8Array
  readonly first: number
  /** The sector size. */
  readonly unit: number
  readonly table: Uint32Array
  /** Sector numbers from this one on are outside (`data`, `table`). */
  readonly limitSectors: number
  readonly sectors: Sectors
  /** Whether a stream in consecutive sectors is a view of `data` (else it is copied). */
  readonly view: boolean
}

/** A stream's sectors, walked as far as its size needs. */
interface Extent {
  readonly medium: Medium
  readonly sectors: ReadonlyArray<number>
  readonly size: number
}

/** Whether each sector follows the one before it. */
const consecutive = (sectors: ReadonlyArray<number>): boolean => {
  // progress: k up to sectors.length
  for (let k = 1; k < sectors.length; k++) if (sectors[k] !== sectors[k - 1]! + 1) return false
  return true
}

/**
 * An extent's bytes, the one read path: a view of the medium when it allows
 * views and the extent is consecutive and inside it, else a copy.
 */
const bytesOf = ({ medium, sectors, size }: Extent): Result.Result<Uint8Array, MalformedFile> => {
  const { data, first, unit } = medium
  const at = sectors.length > 0 ? first + sectors[0]! * unit : 0
  if (medium.view && sectors.length * unit >= size && at + size <= data.length && consecutive(sectors)) {
    // a stream in consecutive sectors (as writers lay out big streams) is a
    // view of the file, not a copy: no pause to copy hundreds of MiB
    return Result.succeed(data.subarray(at, at + size))
  }
  const out = new Uint8Array(size)
  let filled = 0
  // progress: one sector per iteration
  for (const sector of sectors) {
    if (filled >= size) break
    const from = first + sector * unit
    // a partial last sector holds what the data has
    const chunk = data.subarray(from, from + unit)
    const take = Math.min(unit, size - filled)
    if (chunk.length < take) return fail("stream runs past the end of the file")
    out.set(chunk.subarray(0, take), filled)
    filled += take
  }
  if (filled < size) return fail("stream chain is shorter than its size")
  return Result.succeed(out)
}

const utf16 = new TextDecoder("utf-16le")

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
  /** FAT sector k describes sectors k * perSector onwards: the FAT sectors that can describe the file's sectors. */
  const fatSectorsNeeded = Math.ceil(fileSectors / perSector)
  const fileTable = new Sectors(fatSectorsNeeded * perSector)

  const sectorU32 = (sector: number, index: number): number => {
    const at = (sector + 1) * sectorSize + index * 4
    return at + 4 <= bytes.length ? view.getUint32(at, true) : FREESECT
  }

  // FAT sector numbers: 109 in the header, the rest in the DIFAT chain.
  // Some writers (pyExcelerator) declare more FAT sectors than the file has
  // (109 in an 8-sector file) and list only the real ones, the rest free.
  // The declared count only bounds how many numbers are collected (at most
  // one per sector of the file), so slots past it are ignored, junk or not;
  // the FAT is the in-use entries listed. A chain that reaches past the
  // loaded FAT fails when it is walked.
  const numFatWanted = Math.min(u32(44), fileSectors)
  const fatSectors: Array<number> = []
  /** Collects FAT sector numbers from `slots` slots, up to a free one (the list ends) or the count wanted; whether more are wanted. */
  const collect = (slots: number, slot: (k: number) => number): boolean => {
    // progress: k up to slots
    for (let k = 0; k < slots && fatSectors.length < numFatWanted; k++) {
      const sector = slot(k)
      if (sector === FREESECT) return false
      fatSectors.push(sector)
    }
    return fatSectors.length < numFatWanted
  }
  /** A DIFAT pointer; a free one ends the chain too. */
  const difatLink = (sector: number) => sector === FREESECT ? ENDOFCHAIN : sector
  const headerWantsMore = collect(HEADER_DIFAT, (k) => u32(76 + 4 * k))
  // a DIFAT sector lists perSector - 1 FAT sectors, then the next DIFAT
  // sector; its sectors go to the meter with the FAT's, once loaded
  const difat = walk(
    headerWantsMore ? difatLink(u32(68)) : ENDOFCHAIN,
    (sector) => collect(perSector - 1, (k) => sectorU32(sector, k)) ? difatLink(sectorU32(sector, perSector - 1)) : ENDOFCHAIN,
    { what: "DIFAT", limitSectors: fileSectors, max: maxChain, declared: u32(72), sectors: fileTable, meter: noMeter }
  )
  if (Result.isFailure(difat)) return Result.fail(difat.failure)
  if (fatSectors.length === 0) return fail("no FAT sectors")
  // the FAT itself. Only the first fatSectorsNeeded FAT sectors can describe
  // sectors the file has; walks reject sectors past the file before using the
  // FAT, so the rest are never consulted and are not read. (pyExcelerator
  // lists FAT sectors past the end of a small file; those must not fail it,
  // while a FAT sector a chain could need and the file lacks still does.)
  const numFat = Math.min(fatSectors.length, fatSectorsNeeded)
  const fat = new Uint32Array(numFat * perSector)
  // progress: one FAT sector per iteration
  for (let k = 0; k < numFat; k++) {
    const sector = fatSectors[k]!
    if (sector > MAXREGSECT || sector >= fileSectors) return fail(`FAT sector ${sector} is outside the file`)
    const at = (sector + 1) * sectorSize
    const words = Math.floor((Math.min(bytes.length, at + sectorSize) - at) / 4)
    // progress: i up to words
    for (let i = 0; i < words; i++) fat[k * perSector + i] = view.getUint32(at + i * 4, true)
    // a partial last sector leaves the rest free
    for (let i = words; i < perSector; i++) fat[k * perSector + i] = FREESECT
  }
  meter.add("sectors", difat.success.length + numFat)
  // progress: one sector per iteration
  for (const sector of fatSectors) fileTable.reserve(sector, FAT_OWNER)
  for (const sector of difat.success) fileTable.reserve(sector, FAT_OWNER)
  const file: Medium = {
    what: "FAT",
    data: bytes,
    first: sectorSize,
    unit: sectorSize,
    table: fat,
    limitSectors: Math.min(fileSectors, fat.length),
    sectors: fileTable,
    view: true
  }

  /** Up to `count` sectors (all when undefined) of the chain from `start` in `medium`, claimed for `owner`. */
  const chain = (medium: Medium, start: number, count: number | undefined, owner: Owner) => {
    return walk(start, medium.table, {
      what: medium.what,
      limitSectors: medium.limitSectors,
      max: maxChain,
      count,
      owner,
      sectors: medium.sectors,
      meter
    })
  }

  /**
   * The extent of a stream (`name` in messages) of `size` from `start` in
   * `medium`, claimed for `owner`: an invalid size fails here, when the
   * stream is read.
   */
  const extent = (
    name: string,
    size: CfbSize,
    start: number,
    medium: Medium,
    owner: Owner
  ): Result.Result<Extent, CfbError> => {
    if (size._tag === "Invalid") return fail(`${name} is larger than the file`)
    const walked = chain(medium, start, Math.ceil(size.bytes / medium.unit), owner)
    if (Result.isFailure(walked)) return Result.fail(walked.failure)
    return Result.succeed({ medium, sectors: walked.success, size: size.bytes })
  }

  // the directory
  const dirChain = chain(file, u32(48), undefined, DIRECTORY_OWNER)
  if (Result.isFailure(dirChain)) return Result.fail(dirChain.failure)
  const dirRead = bytesOf({ medium: file, sectors: dirChain.success, size: dirChain.success.length * sectorSize })
  if (Result.isFailure(dirRead)) return Result.fail(dirRead.failure)
  const dirBytes = dirRead.success
  const dirView = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength)
  const count = Math.floor(dirBytes.length / DIRECTORY_ENTRY)
  if (count === 0) return fail("empty directory")

  interface Raw {
    readonly name: string
    readonly type: number
    readonly left: number
    readonly right: number
    readonly child: number
    readonly start: number
    readonly size: CfbSize
  }
  const raw = (id: number): Result.Result<Raw, MalformedFile> => {
    const at = id * DIRECTORY_ENTRY
    const nameLength = dirView.getUint16(at + 64, true)
    if (nameLength > 64 || nameLength % 2 !== 0) return fail(`directory entry ${id} has a bad name length`)
    const name = utf16.decode(dirBytes.subarray(at, at + Math.max(0, nameLength - 2)))
    const low = dirView.getUint32(at + 120, true)
    const high = dirView.getUint32(at + 124, true)
    // version 3 files may leave garbage in the high word; a size past 2^53
    // is not exact, but larger than any file all the same
    const declared = major === 3 ? low : high * 0x100000000 + low
    return Result.succeed({
      name,
      type: dirView.getUint8(at + 66),
      left: dirView.getUint32(at + 68, true),
      right: dirView.getUint32(at + 72, true),
      child: dirView.getUint32(at + 76, true),
      start: dirView.getUint32(at + 116, true),
      size: declared > fileSectors * sectorSize ? { _tag: "Invalid" } : { _tag: "Valid", bytes: declared }
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
    const size: CfbSize = type === "stream" ? entry.size : { _tag: "Valid", bytes: 0 }
    entries.push({ id, name: entry.name, path, type, size, start: entry.start })
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

  // the mini stream (the root entry's stream) and the miniFAT, read on first use
  let mini: Result.Result<Medium, CfbError> | undefined
  const loadMini = (): Result.Result<Medium, CfbError> => {
    if (mini !== undefined) return mini
    const numMiniFat = u32(64)
    const miniFatChain = chain(file, u32(60), numMiniFat, MINIFAT_OWNER)
    if (Result.isFailure(miniFatChain)) return (mini = Result.fail(miniFatChain.failure))
    const miniFatSectors = miniFatChain.success
    if (miniFatSectors.length < numMiniFat) return (mini = fail("miniFAT chain is shorter than the header says"))
    const table = new Uint32Array(miniFatSectors.length * perSector)
    // progress: one sector per iteration, i up to perSector
    miniFatSectors.forEach((sector, k) => {
      for (let i = 0; i < perSector; i++) table[k * perSector + i] = sectorU32(sector, i)
    })
    const rootExtent = extent("mini stream", root.size, root.start, file, streamOwner(0))
    if (Result.isFailure(rootExtent)) return (mini = Result.fail(rootExtent.failure))
    // a view of the caller's buffer when consecutive: mini streams are
    // copied out of it when read (see `Cfb.read`)
    const data = bytesOf(rootExtent.success)
    if (Result.isFailure(data)) return (mini = Result.fail(data.failure))
    // a last, partial mini sector counts (its stream reads only what it has)
    const miniSectors = Math.ceil(data.success.length / MINI_SECTOR)
    return (mini = Result.succeed({
      what: "miniFAT",
      data: data.success,
      first: 0,
      unit: MINI_SECTOR,
      table,
      limitSectors: Math.min(miniSectors, table.length),
      sectors: new Sectors(table.length),
      view: false
    }))
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
      // an empty stream reads without loading the mini stream
      if (size._tag === "Valid" && size.bytes === 0) return Result.succeed(new Uint8Array(0))
      let medium = file
      if (size._tag === "Valid" && size.bytes < MINI_CUTOFF) {
        const m = loadMini()
        if (Result.isFailure(m)) return Result.fail(m.failure)
        medium = m.success
      }
      const read = extent(`stream ${entry.path}`, size, entry.start, medium, streamOwner(entry.id))
      return Result.isFailure(read) ? Result.fail(read.failure) : bytesOf(read.success)
    }
  })
}

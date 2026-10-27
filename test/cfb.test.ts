import { readdirSync, readFileSync, statSync } from "node:fs"
import { afterAll, describe, expect, it } from "vitest"
import { FuzzHost } from "./fuzz/host.ts"
import { RecordReader, readBof, EOF } from "../src/internal/biff/records.ts"
import { type Cfb, type CfbEntry, open } from "../src/internal/cfb/cfb.ts"
import { MalformedFile } from "../src/SpreadsheetError.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import { type CfbItem, patch32, sectorOffset, writeCfb } from "./helpers/cfb.ts"

const limits = SpreadsheetLimits.defaults

const pseudoRandom = (n: number, seed = 1): Uint8Array => {
  const out = new Uint8Array(n)
  let x = seed
  for (let i = 0; i < n; i++) {
    x = (x * 1103515245 + 12345) >>> 0
    out[i] = x >>> 24
  }
  return out
}

const openOk = (bytes: Uint8Array, l = limits): Cfb => {
  const cfb = open(bytes, l)
  if ("_tag" in cfb) throw new Error(`open failed: ${cfb._tag} ${(cfb as { reason?: string }).reason}`)
  return cfb
}

const readOk = (cfb: Cfb, path: string): Uint8Array => {
  const data = cfb.read(cfb.find(path)!)
  if ("_tag" in data) throw new Error(`read failed: ${data._tag} ${(data as { reason?: string }).reason}`)
  return data
}

const reasonOf = (result: unknown): string =>
  (result as { reason?: string; _tag?: string }).reason ?? (result as { _tag?: string })._tag ?? "ok"

const SAMPLE: ReadonlyArray<CfbItem> = [
  { name: "Workbook", data: pseudoRandom(10_000, 1) },
  { name: "\u0005SummaryInformation", data: pseudoRandom(300, 2) },
  { name: "Empty", data: new Uint8Array(0) },
  { name: "Edge", data: pseudoRandom(4095, 3) },
  { name: "Exact", data: pseudoRandom(4096, 4) },
  { name: "_VBA_PROJECT_CUR", children: [{ name: "VBA", children: [{ name: "dir", data: pseudoRandom(700, 5) }] }] }
]

describe("cfb: reading", () => {
  it("reads big, mini, empty and nested streams back exactly", () => {
    const { bytes } = writeCfb(SAMPLE)
    const cfb = openOk(bytes)
    expect(cfb.entries.map((e) => [e.path, e.type])).toEqual([
      ["Workbook", "stream"],
      ["\u0005SummaryInformation", "stream"],
      ["Empty", "stream"],
      ["Edge", "stream"],
      ["Exact", "stream"],
      ["_VBA_PROJECT_CUR", "storage"],
      ["_VBA_PROJECT_CUR/VBA", "storage"],
      ["_VBA_PROJECT_CUR/VBA/dir", "stream"]
    ])
    expect(readOk(cfb, "workbook")).toEqual(pseudoRandom(10_000, 1))
    expect(readOk(cfb, "\u0005SummaryInformation")).toEqual(pseudoRandom(300, 2))
    expect(readOk(cfb, "Empty")).toEqual(new Uint8Array(0))
    expect(readOk(cfb, "Edge")).toEqual(pseudoRandom(4095, 3))
    expect(readOk(cfb, "Exact")).toEqual(pseudoRandom(4096, 4))
    expect(readOk(cfb, "_vba_project_cur/vba/DIR")).toEqual(pseudoRandom(700, 5))
    expect(reasonOf(cfb.read(cfb.find("_VBA_PROJECT_CUR")!))).toMatch(/not a stream/)
    // work: every sector at most a few times
    expect(cfb.sectorsVisited()).toBeLessThanOrEqual(3 * (bytes.length / 512))
  })

  it("a stream in consecutive sectors is a view of the file; a fragmented one is gathered", () => {
    const layout = writeCfb(SAMPLE)
    const contiguous = openOk(layout.bytes)
    const view = readOk(contiguous, "Workbook")
    expect(view).toEqual(pseudoRandom(10_000, 1))
    expect(view.buffer).toBe(layout.bytes.buffer)
    // swap Workbook's 2nd and 3rd sectors, data and FAT links: same stream, out of order
    const s = layout.streams.get("Workbook")!.start
    const bytes = layout.bytes.slice()
    const second = bytes.slice(sectorOffset(s + 1), sectorOffset(s + 2))
    bytes.copyWithin(sectorOffset(s + 1), sectorOffset(s + 2), sectorOffset(s + 3))
    bytes.set(second, sectorOffset(s + 2))
    const fat = sectorOffset(layout.fatSectors[0]!)
    patch32(bytes, fat + 4 * s, s + 2)
    patch32(bytes, fat + 4 * (s + 2), s + 1)
    patch32(bytes, fat + 4 * (s + 1), s + 3)
    const fragmented = readOk(openOk(bytes), "Workbook")
    expect(fragmented).toEqual(pseudoRandom(10_000, 1))
    expect(fragmented.buffer).not.toBe(bytes.buffer)
  })

  it("reads version 4 files (4096-byte sectors), with 64-bit sizes", () => {
    const layout = writeCfb(SAMPLE, { version: 4 })
    expect(layout.sectorSize).toBe(4096)
    const cfb = openOk(layout.bytes)
    expect(readOk(cfb, "Workbook")).toEqual(pseudoRandom(10_000, 1))
    expect(readOk(cfb, "_VBA_PROJECT_CUR/VBA/dir")).toEqual(pseudoRandom(700, 5))
    // a size's high word counts in version 4
    const dir = sectorOffset(layout.directorySectors[0]!, 4096) + layout.streams.get("Workbook")!.id * 128
    for (const high of [1, 0x200000]) {
      const big = layout.bytes.slice()
      patch32(big, dir + 124, high)
      const c = openOk(big)
      expect(reasonOf(c.read(c.find("Workbook")!)), String(high)).toMatch(/larger than the file/)
    }
    // version 4 needs 4096-byte sectors, version 3 512-byte ones
    const wrongShift = layout.bytes.slice()
    wrongShift[30] = 9
    expect(reasonOf(open(wrongShift, limits))).toBe("unsupported version 4 with sector size 2^9")
  })

  it("version 3 ignores the high word of a stream size (writers leave garbage there)", () => {
    const layout = writeCfb(SAMPLE)
    const garbage = layout.bytes.slice()
    patch32(garbage, sectorOffset(layout.directorySectors[0]!) + layout.streams.get("Workbook")!.id * 128 + 124, 0xDEADBEEF)
    expect(readOk(openOk(garbage), "Workbook")).toEqual(pseudoRandom(10_000, 1))
  })

  it("a FAT count larger than the file (as pyExcelerator writes: 109 in a small file) reads", () => {
    const layout = writeCfb(SAMPLE)
    expect(layout.fatSectors.length).toBe(1)
    const lying = layout.bytes.slice()
    patch32(lying, 44, 109)
    expect(109).toBeGreaterThan(lying.length / 512 - 1)
    const cfb = openOk(lying)
    expect(readOk(cfb, "Workbook")).toEqual(pseudoRandom(10_000, 1))
    expect(readOk(cfb, "Edge")).toEqual(pseudoRandom(4095, 3))
    // a count within the file that the listed sectors don't reach is still MalformedFile
    const short = layout.bytes.slice()
    patch32(short, 44, 3)
    expect(reasonOf(open(short, limits))).toBe("found 1 of 3 FAT sectors")
    const none = layout.bytes.slice()
    patch32(none, 76, 0xFFFFFFFF)
    expect(reasonOf(open(none, limits))).toBe("no FAT sectors")
  })

  it("FAT sectors listed past the end of the file that only describe missing sectors are not read (as pyExcelerator writes)", () => {
    // 109 FAT sectors declared and listed, the first the real one (the last
    // sector of the file), the other 108 past the end of the file
    const layout = writeCfb(SAMPLE)
    const fat = layout.fatSectors[0]!
    expect(layout.fatSectors.length).toBe(1)
    expect(fat).toBe(layout.bytes.length / 512 - 2)
    const listed = layout.bytes.slice()
    patch32(listed, 44, 109)
    for (let k = 1; k < 109; k++) patch32(listed, 76 + 4 * k, fat + k)
    const cfb = openOk(listed)
    expect(readOk(cfb, "Workbook")).toEqual(pseudoRandom(10_000, 1))
    expect(readOk(cfb, "Edge")).toEqual(pseudoRandom(4095, 3))
    expect(readOk(cfb, "_VBA_PROJECT_CUR/VBA/dir")).toEqual(pseudoRandom(700, 5))
    // a FAT sector that describes sectors the file has, and is missing, still fails
    const big = writeCfb([{ name: "Workbook", data: pseudoRandom(200 * 512, 9) }])
    expect(big.fatSectors.length).toBe(2)
    const missing = big.bytes.slice()
    patch32(missing, 80, big.bytes.length / 512 + 5)
    expect(reasonOf(open(missing, limits))).toBe(`FAT sector ${big.bytes.length / 512 + 5} is outside the file`)
  })

  it("a lying FAT count stops collecting FAT sectors at the file's sector count (no DIFAT walk past it)", () => {
    // a huge declared count, all 109 header entries listed (the real FAT
    // sector, then past the end of the file) and a DIFAT start that repeats
    // the FAT sector (no DIFAT sectors declared) or points past the file:
    // only the file's sectors' worth of FAT sectors are wanted, so the
    // header list is enough and the bogus DIFAT is never followed
    const layout = writeCfb(SAMPLE)
    const fat = layout.fatSectors[0]!
    const fileSectors = layout.bytes.length / 512 - 1
    expect(fileSectors).toBeLessThan(109)
    for (const difatStart of [fat, fileSectors + 50]) {
      const lying = layout.bytes.slice()
      patch32(lying, 44, 1_000_000)
      for (let k = 1; k < 109; k++) patch32(lying, 76 + 4 * k, fat + k)
      patch32(lying, 68, difatStart)
      patch32(lying, 72, 0)
      const cfb = openOk(lying)
      expect(readOk(cfb, "Workbook")).toEqual(pseudoRandom(10_000, 1))
      expect(readOk(cfb, "Edge")).toEqual(pseudoRandom(4095, 3))
    }
  })

  it("a partial last sector is read as far as it goes (pinned: the file opens and reads)", () => {
    const layout = writeCfb(SAMPLE)
    // the last sector is the FAT; its last 25 entries (unused) are cut off
    expect(layout.fatSectors.at(-1)).toBe(layout.bytes.length / 512 - 2)
    const cut = layout.bytes.subarray(0, layout.bytes.length - 100)
    const cfb = openOk(cut)
    expect(readOk(cfb, "Workbook")).toEqual(pseudoRandom(10_000, 1))
    expect(readOk(cfb, "Edge")).toEqual(pseudoRandom(4095, 3))
  })

  it("reads a file whose FAT needs a DIFAT (more than 109 FAT sectors)", () => {
    const big = pseudoRandom(110 * 128 * 512, 7)
    const layout = writeCfb([{ name: "Workbook", data: big }])
    expect(layout.difatSectors.length).toBeGreaterThan(0)
    const cfb = openOk(layout.bytes)
    expect(Buffer.compare(readOk(cfb, "Workbook"), big)).toBe(0)
    // the DIFAT walk is bounded by maxSectorChain too: a 2-sector DIFAT
    // under a limit of 1 (the directory chain is 1 sector, so only the
    // DIFAT can hit it)
    const twoDifat = writeCfb([{ name: "Workbook", data: new Uint8Array(240 * 128 * 512) }])
    expect(twoDifat.difatSectors.length).toBe(2)
    expect(twoDifat.directorySectors.length).toBe(1)
    expect(open(twoDifat.bytes, { ...limits, maxSectorChain: 1 })).toMatchObject({ _tag: "LimitExceeded", limit: "maxSectorChain", max: 1, actual: 2 })
  })

  it("opens every committed .xls in a CFB, and its workbook stream is a BIFF record stream from BOF to EOF", () => {
    let files = 0
    for (const dir of ["test/corpus/generated", "test/corpus/hidden-rows", "test/corpus/hidden-rows/renamed"]) {
      for (const name of readdirSync(dir)) {
        if (!statSync(`${dir}/${name}`).isFile()) continue
        const bytes = new Uint8Array(readFileSync(`${dir}/${name}`))
        if (bytes[0] !== 0xD0 || bytes[1] !== 0xCF) continue
        const cfb = openOk(bytes)
        const entry: CfbEntry | undefined = cfb.find("Workbook") ?? cfb.find("Book")
        expect(entry, name).toBeDefined()
        const stream = readOk(cfb, entry!.path)
        const reader = new RecordReader(stream, entry!.path)
        const first = reader.next()
        expect(first instanceof MalformedFile || first === undefined ? undefined : readBof(first)?.kind, name).toBe(0x0005)
        let eofs = 0
        let record = reader.next()
        // progress: each next() advances the offset
        while (record !== undefined) {
          if (record instanceof MalformedFile) throw new Error(`${name}: ${record.reason}`)
          if (record.type === EOF) eofs++
          record = reader.next()
        }
        expect(eofs, name).toBeGreaterThanOrEqual(2) // globals and at least one sheet
        expect(reader.records, name).toBeLessThanOrEqual(stream.length / 4)
        for (const e of cfb.entries.filter((x) => x.type === "stream")) expect("_tag" in cfb.read(e), `${name}: ${e.path}`).toBe(false)
        files++
      }
    }
    expect(files).toBeGreaterThanOrEqual(7)
  })
})

describe("cfb: hostile and broken files", () => {
  const fresh = () => writeCfb(SAMPLE)
  /** The FAT entry for `sector` (all test files have one FAT sector here). */
  const fatEntry = (layout: ReturnType<typeof writeCfb>, sector: number) => sectorOffset(layout.fatSectors[0]!) + 4 * sector

  it("bad headers are MalformedFile", () => {
    const cases: Array<[string, (b: Uint8Array) => void]> = [
      ["not a CFB file (bad signature)", (b) => void (b[0] = 0)],
      ["bad byte order mark", (b) => void (b[28] = 0)],
      ["unsupported version 3 with sector size 2^12", (b) => void (b[30] = 12)],
      ["mini sector size must be 64", (b) => void (b[32] = 7)],
      ["mini stream cutoff must be 4096", (b) => patch32(b, 56, 8192)]
    ]
    for (const [reason, mutate] of cases) {
      const { bytes } = fresh()
      mutate(bytes)
      expect(reasonOf(open(bytes, limits)), reason).toBe(reason)
    }
    expect(reasonOf(open(new Uint8Array(100), limits))).toBe("file shorter than a CFB header")
  })

  it("a FAT chain with a cycle, or pointing outside the file, is MalformedFile", () => {
    const layout = fresh()
    const start = layout.streams.get("Workbook")!.start
    const cycle = layout.bytes.slice()
    patch32(cycle, fatEntry(layout, start + 3), start + 1) // ... -> 3 -> 1 -> 2 -> 3
    const a = openOk(cycle)
    expect(reasonOf(a.read(a.find("Workbook")!))).toMatch(/FAT chain has a cycle/)
    const outside = layout.bytes.slice()
    patch32(outside, fatEntry(layout, start), 100_000)
    const b = openOk(outside)
    expect(reasonOf(b.read(b.find("Workbook")!))).toMatch(/FAT chain points at sector 100000/)
    // inside the FAT table (128 entries) but past the file's sectors
    const pastFile = layout.bytes.slice()
    expect(pastFile.length / 512 - 1).toBeLessThan(100)
    patch32(pastFile, fatEntry(layout, start), 100)
    const d = openOk(pastFile)
    expect(reasonOf(d.read(d.find("Workbook")!))).toBe("FAT chain points at sector 100")
    const early = layout.bytes.slice()
    patch32(early, fatEntry(layout, start + 5), 0xFFFFFFFE) // ends after 6 of 20 sectors
    const c = openOk(early)
    expect(reasonOf(c.read(c.find("Workbook")!))).toBe("stream chain is shorter than its size")
  })

  it("stream sizes larger than the file, and long chains past maxSectorChain, stop", () => {
    const layout = fresh()
    const huge = layout.bytes.slice()
    patch32(huge, sectorOffset(layout.directorySectors[0]!) + layout.streams.get("Workbook")!.id * 128 + 120, 0x7FFFFFFF)
    const a = openOk(huge)
    expect(reasonOf(a.read(a.find("Workbook")!))).toMatch(/larger than the file/)
    const b = openOk(layout.bytes, { ...limits, maxSectorChain: 5 })
    expect(b.read(b.find("Workbook")!)).toMatchObject({ _tag: "LimitExceeded", limit: "maxSectorChain", max: 5 })
  })

  // Run in the fuzz worker with a time budget: if a loop check were
  // removed, the walk would never end; the worker is then killed and the
  // test fails naming the hang instead of freezing the run.
  const host = new FuzzHost(5000)
  afterAll(() => host.stop())

  it("directory loops and bad entries are MalformedFile (run in a worker with a budget)", async () => {
    const layout = fresh()
    const dir = sectorOffset(layout.directorySectors[0]!)
    const cases: Array<[string, (b: Uint8Array) => void]> = [
      ["directory entry 1 is reached twice", (b) => patch32(b, dir + 1 * 128 + 72, 1)], // its own right sibling
      ["directory entry 0 is reached twice", (b) => patch32(b, dir + 2 * 128 + 72, 0)], // a sibling pointing at the root
      ["directory entry 6 is reached twice", (b) => patch32(b, dir + 7 * 128 + 76, 6)], // a storage whose child is its parent
      ["directory points at entry 5000 of 12", (b) => patch32(b, dir + 1 * 128 + 72, 5000)],
      ["directory entry 1 has a bad name length", (b) => void (b[dir + 1 * 128 + 64] = 99)],
      ["directory entry 1 has type 9", (b) => void (b[dir + 1 * 128 + 66] = 9)],
      ["the first directory entry is not the root", (b) => void (b[dir + 66] = 1)],
      ['directory entry 1 has "/" in its name', (b) => void (b[dir + 1 * 128 + 2] = 0x2F)]
    ]
    for (const [reason, mutate] of cases) {
      const bytes = layout.bytes.slice()
      mutate(bytes)
      expect(await host.run({ op: "cfb", bytes }), reason).toEqual({ kind: "typed", tag: "MalformedFile", reason })
    }
  }, 120_000)

  it("a directory chain cycle, a miniFAT cycle and a bad mini stream are MalformedFile", () => {
    const layout = fresh()
    const dirCycle = layout.bytes.slice()
    patch32(dirCycle, fatEntry(layout, layout.directorySectors.at(-1)!), layout.directorySectors[0]!)
    expect(reasonOf(open(dirCycle, limits))).toMatch(/FAT chain has a cycle/)
    const miniCycle = layout.bytes.slice()
    const s = layout.streams.get("\u0005SummaryInformation")!.start
    patch32(miniCycle, sectorOffset(layout.miniFatSectors[0]!) + 4 * (s + 1), s)
    const a = openOk(miniCycle)
    expect(reasonOf(a.read(a.find("\u0005SummaryInformation")!))).toMatch(/miniFAT chain has a cycle/)
    const shortMiniFat = layout.bytes.slice()
    patch32(shortMiniFat, 64, 5)
    const d = openOk(shortMiniFat)
    expect(reasonOf(d.read(d.find("Edge")!))).toBe("miniFAT chain is shorter than the header says")
    const bigRoot = layout.bytes.slice()
    patch32(bigRoot, sectorOffset(layout.directorySectors[0]!) + 120, 0x7FFFFFF0)
    const b = openOk(bigRoot)
    expect(reasonOf(b.read(b.find("Edge")!))).toBe("mini stream is larger than the file")
  })

  it("a DIFAT chain with a cycle, or longer than the header says, is MalformedFile", () => {
    const layout = writeCfb([{ name: "Workbook", data: new Uint8Array(110 * 128 * 512) }])
    const difat = layout.difatSectors[0]!
    const cycle = layout.bytes.slice()
    patch32(cycle, 72, 5) // allow more DIFAT sectors
    patch32(cycle, 44, 109 + 5 * 127) // and ask for more FAT sectors than exist in one
    // a full DIFAT sector (its free entries filled in), whose next pointer is itself
    for (let k = 2; k < 127; k++) patch32(cycle, sectorOffset(difat) + 4 * k, layout.fatSectors[0]!)
    patch32(cycle, sectorOffset(difat) + 508, difat)
    const result = open(cycle, limits)
    expect(reasonOf(result)).toBe(`DIFAT chain has a cycle at sector ${difat}`)
    const tooLong = layout.bytes.slice()
    patch32(tooLong, 44, 109 + 2 * 127)
    for (let k = 2; k < 127; k++) patch32(tooLong, sectorOffset(difat) + 4 * k, layout.fatSectors[0]!)
    patch32(tooLong, sectorOffset(difat) + 508, layout.fatSectors[0]!)
    expect(reasonOf(open(tooLong, limits))).toBe("DIFAT chain is longer than the header says")
    // a FAT sector number outside the file
    const outside = layout.bytes.slice()
    patch32(outside, 76, 99_999)
    expect(reasonOf(open(outside, limits))).toBe("FAT sector 99999 is outside the file")
    // a FAT count beyond what the file can hold: the listed in-use FAT sectors are the FAT
    const tooMany = layout.bytes.slice()
    patch32(tooMany, 44, 1_000_000)
    expect(Buffer.compare(readOk(openOk(tooMany), "Workbook"), new Uint8Array(110 * 128 * 512))).toBe(0)
  })

  it("streams that share sectors are MalformedFile; reading one stream twice is fine", () => {
    const layout = writeCfb([...SAMPLE, { name: "Second", data: pseudoRandom(10_000, 8) }])
    const dir = sectorOffset(layout.directorySectors[0]!)
    const startOf = (path: string) => layout.streams.get(path)!
    // Second's directory entry pointed at Workbook's sectors
    const shared = layout.bytes.slice()
    patch32(shared, dir + startOf("Second").id * 128 + 116, startOf("Workbook").start)
    const a = openOk(shared)
    expect(readOk(a, "Workbook")).toEqual(pseudoRandom(10_000, 1))
    expect(readOk(a, "Workbook")).toEqual(pseudoRandom(10_000, 1))
    expect(reasonOf(a.read(a.find("Second")!))).toMatch(/belongs to two streams/)
    // in the mini stream too
    const sharedMini = layout.bytes.slice()
    patch32(sharedMini, dir + startOf("Edge").id * 128 + 116, startOf("\u0005SummaryInformation").start)
    const b = openOk(sharedMini)
    readOk(b, "\u0005SummaryInformation")
    expect(reasonOf(b.read(b.find("Edge")!))).toMatch(/miniFAT belongs to two streams/)
    // a stream running into the directory or a FAT sector
    for (const sector of [layout.directorySectors[0]!, layout.fatSectors[0]!]) {
      const intoOwn = layout.bytes.slice()
      patch32(intoOwn, dir + startOf("Second").id * 128 + 116, sector)
      const c = openOk(intoOwn)
      expect(reasonOf(c.read(c.find("Second")!)), String(sector)).toMatch(/belongs to two streams/)
    }
  })

  it("duplicate paths: the lower directory id wins; mini streams ending mid sector read", () => {
    const { bytes } = writeCfb([{ name: "Same", data: new Uint8Array([1]) }, { name: "SAME", data: new Uint8Array([2]) }, { name: "Odd", data: pseudoRandom(100, 3) }])
    const cfb = openOk(bytes)
    expect(cfb.find("same")!.id).toBe(1)
    expect(readOk(cfb, "Same")).toEqual(new Uint8Array([1]))
    // the mini stream ends inside its last mini sector: its size cut to end there
    const layout = writeCfb([{ name: "Odd", data: pseudoRandom(100, 3) }])
    const trimmed = layout.bytes.slice()
    patch32(trimmed, sectorOffset(layout.directorySectors[0]!) + 120, 100) // root (mini stream) size 100, not 128
    expect(readOk(openOk(trimmed), "Odd")).toEqual(pseudoRandom(100, 3))
  })

  it("a stream starting at a DIFAT sector claims a sector the container owns", () => {
    const layout = writeCfb([{ name: "Workbook", data: new Uint8Array(110 * 128 * 512) }, { name: "Other", data: new Uint8Array(5000) }])
    expect(layout.difatSectors.length).toBe(1)
    const bytes = layout.bytes.slice()
    patch32(bytes, sectorOffset(layout.directorySectors[0]!) + layout.streams.get("Other")!.id * 128 + 116, layout.difatSectors[0]!)
    const cfb = openOk(bytes)
    expect(reasonOf(cfb.read(cfb.find("Other")!))).toMatch(/belongs to two streams/)
  })

  it("deeply nested storages stop at maxStorageDepth (paths would cost depth²)", () => {
    let item: CfbItem = { name: "leaf", data: new Uint8Array([1]) }
    for (let k = 0; k < 10_000; k++) item = { name: `s${k % 10}`, children: [item] }
    const { bytes } = writeCfb([item])
    expect(open(bytes, limits)).toMatchObject({ _tag: "LimitExceeded", limit: "maxStorageDepth", max: 32, actual: 33 })
    // within the limit it reads
    let shallow: CfbItem = { name: "leaf", data: new Uint8Array([7]) }
    for (let k = 0; k < 31; k++) shallow = { name: "s", children: [shallow] }
    const cfb = openOk(writeCfb([shallow]).bytes)
    expect(readOk(cfb, `${"s/".repeat(31)}leaf`)).toEqual(new Uint8Array([7]))
  })

  it("truncated files fail with a typed error, never a throw", () => {
    const { bytes } = writeCfb([...SAMPLE, { name: "Big", data: pseudoRandom(50_000, 9) }])
    for (let cut = 0; cut < bytes.length; cut += 97) {
      const cfb = open(bytes.subarray(0, cut), limits)
      if ("_tag" in cfb) {
        expect(cfb._tag).toBe("MalformedFile")
        continue
      }
      for (const entry of cfb.entries.filter((e) => e.type === "stream")) {
        const data = cfb.read(entry)
        if ("_tag" in data) expect(data._tag).toBe("MalformedFile")
      }
    }
  })
})

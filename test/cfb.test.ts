import { readdirSync, readFileSync, statSync } from "node:fs"
import { describe, expect, it } from "vitest"
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

  it("reads a file whose FAT needs a DIFAT (more than 109 FAT sectors)", () => {
    const big = pseudoRandom(110 * 128 * 512, 7)
    const layout = writeCfb([{ name: "Workbook", data: big }])
    expect(layout.difatSectors.length).toBeGreaterThan(0)
    const cfb = openOk(layout.bytes)
    expect(readOk(cfb, "Workbook")).toEqual(big)
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

  it("directory loops and bad entries are MalformedFile", () => {
    const layout = fresh()
    const dir = sectorOffset(layout.directorySectors[0]!)
    const selfSibling = layout.bytes.slice()
    patch32(selfSibling, dir + 1 * 128 + 72, 1) // entry 1's right sibling is itself
    expect(reasonOf(open(selfSibling, limits))).toBe("directory entry 1 is reached twice")
    const toRoot = layout.bytes.slice()
    patch32(toRoot, dir + 2 * 128 + 72, 0) // a sibling pointing at the root
    expect(reasonOf(open(toRoot, limits))).toBe("directory entry 0 is reached twice")
    const outside = layout.bytes.slice()
    patch32(outside, dir + 1 * 128 + 72, 5000)
    expect(reasonOf(open(outside, limits))).toMatch(/directory points at entry 5000/)
    const badName = layout.bytes.slice()
    badName[dir + 1 * 128 + 64] = 99
    expect(reasonOf(open(badName, limits))).toBe("directory entry 1 has a bad name length")
    const badType = layout.bytes.slice()
    badType[dir + 1 * 128 + 66] = 9
    expect(reasonOf(open(badType, limits))).toBe("directory entry 1 has type 9")
    const noRoot = layout.bytes.slice()
    noRoot[dir + 66] = 1
    expect(reasonOf(open(noRoot, limits))).toBe("the first directory entry is not the root")
  })

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
    patch32(cycle, sectorOffset(difat) + 508, difat)
    const result = open(cycle, limits)
    expect(reasonOf(result)).toBe(`DIFAT chain has a cycle at sector ${difat}`)
    const tooLong = layout.bytes.slice()
    patch32(tooLong, 44, 109 + 2 * 127)
    patch32(tooLong, sectorOffset(difat) + 508, layout.fatSectors[0]!)
    expect(reasonOf(open(tooLong, limits))).toBe("DIFAT chain is longer than the header says")
    // a FAT count beyond what the file can hold
    const tooMany = layout.bytes.slice()
    patch32(tooMany, 44, 1_000_000)
    expect(reasonOf(open(tooMany, limits))).toMatch(/header lists 1000000 FAT sectors/)
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

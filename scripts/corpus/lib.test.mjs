// Tests for scripts/corpus/lib.mjs helpers. Run: node --test scripts/corpus/lib.test.mjs
import assert from "node:assert/strict"
import fs from "node:fs"
import os from "node:os"
import path from "node:path"
import test from "node:test"
import { GENERATED_DIR, biffVersion, sniff } from "./lib.mjs"

const bof = (type, vers, size, dt = 0x0005) => {
  const b = Buffer.alloc(4 + Math.max(size, 4) + 4)
  b.writeUInt16LE(type, 0); b.writeUInt16LE(size, 2); b.writeUInt16LE(vers, 4); b.writeUInt16LE(dt, 6)
  return b
}

test("biffVersion: BIFF8 BOF per [MS-XLS] 2.4.21 (0x0809, vers 0x0600, 16 bytes)", () => {
  assert.equal(biffVersion(bof(0x0809, 0x0600, 16)), 8)
  assert.equal(biffVersion(bof(0x0809, 0x0600, 16, 0x0010)), 8) // worksheet substream
})

test("biffVersion: BIFF5 (0x0809, vers 0x0500, 8 bytes)", () => {
  assert.equal(biffVersion(bof(0x0809, 0x0500, 8)), 5)
})

test("biffVersion: other vers under 0x0809 go by record size", () => {
  assert.equal(biffVersion(bof(0x0809, 0x1234, 16)), 8)
  assert.equal(biffVersion(bof(0x0809, 0x1234, 8)), 5)
})

test("biffVersion: BIFF2/3/4 BOF record types", () => {
  assert.equal(biffVersion(bof(0x0009, 0x0002, 4)), 2)
  assert.equal(biffVersion(bof(0x0209, 0x0003, 6)), 3)
  assert.equal(biffVersion(bof(0x0409, 0x0004, 6)), 4)
})

test("biffVersion: not a BOF", () => {
  assert.equal(biffVersion(Buffer.alloc(4)), null) // too short
  assert.equal(biffVersion(bof(0x0208, 0x0600, 16)), null) // ROW, not BOF
  assert.equal(biffVersion(bof(0x0809, 0x0600, 2)), null) // record too small for vers + dt
  assert.equal(biffVersion(bof(0x0809, 0x0600, 21)), null) // longer than any BOF
  assert.equal(biffVersion(Buffer.from("PK\x03\x04 not biff")), null)
})

test("biffVersion via sniff on the generated xls fixtures", () => {
  const expect = { "basic.biff2.xls": 2, "basic.biff3.xls": 3, "basic.biff4.xls": 4, "features.biff5.xls": 5, "hidden_rows.biff5.xls": 5, "features.xls": 8, "hidden_rows.xls": 8, "sparse_far.xls": 8, "huge_dimension.xls": 8 }
  for (const [name, biff] of Object.entries(expect)) {
    const s = sniff(fs.readFileSync(path.join(GENERATED_DIR, name)))
    assert.equal(s.format, "xls", name)
    assert.equal(s.features.biff, biff, name)
  }
})

test("sameZipContent: equal content with other deflate bytes, not when content or order differs", async () => {
  const { zip, sameZipContent } = await import("./lib.mjs")
  const parts = [["a.xml", Buffer.from("<a>" + "xy".repeat(3000) + "<b/>".repeat(500) + "</a>")], ["b.txt", Buffer.from("hello")]]
  const ours = zip(parts)
  const relevelled = zip(parts, { level: 1 }) // stands in for another zlib build
  assert.ok(!ours.equals(relevelled), "different deflate bytes")
  assert.ok(sameZipContent(ours, relevelled))
  assert.ok(!sameZipContent(ours, zip([parts[0], ["b.txt", Buffer.from("hellO")]])), "content change")
  assert.ok(!sameZipContent(ours, zip([parts[1], parts[0]])), "order change")
  assert.ok(!sameZipContent(ours, zip([parts[0]])), "missing entry")
  assert.ok(!sameZipContent(Buffer.from("not a zip"), ours))
})

// ---- zipDiff: only compressed bytes may differ (method, compressed size, data, offsets)
async function zipFixtures() {
  const { zip, zipDiff, sameZipContent, compareDirs, zipEntries } = await import("./lib.mjs")
  const parts = [["xl/a.xml", Buffer.from("<a>" + "xy".repeat(3000) + "</a>")], ["b.txt", Buffer.from("hello")]]
  const base = zip(parts)
  const offsets = (buf) => {
    const ents = zipEntries(buf)
    const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
    const cd = buf.readUInt32LE(eocd + 16)
    const central = []
    for (let p = cd, i = 0; i < ents.length; i++) { central.push(p); p += 46 + buf.readUInt16LE(p + 28) + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32) }
    return { local: ents.map((e) => e.lho), central, eocd }
  }
  return { zip, zipDiff, sameZipContent, compareDirs, parts, base, offsets }
}

test("zipDiff: recompressed zip (other deflate level) is the same", async () => {
  const { zip, zipDiff, parts, base } = await zipFixtures()
  assert.equal(zipDiff(base, zip(parts, { level: 1 })), null)
})

test("zipDiff: changed DOS timestamps in the headers are a difference", async () => {
  const { zipDiff, sameZipContent, base, offsets } = await zipFixtures()
  const t = Buffer.from(base), o = offsets(t)
  for (const p of o.local) { t.writeUInt16LE(0x6000, p + 10); t.writeUInt16LE(0x5a21, p + 12) }
  for (const p of o.central) { t.writeUInt16LE(0x6000, p + 12); t.writeUInt16LE(0x5a21, p + 14) }
  assert.match(zipDiff(base, t), /header fields/)
  assert.equal(sameZipContent(base, t), false)
  // central only, or local only, is caught as well
  const c = Buffer.from(base); c.writeUInt16LE(0x5a21, offsets(c).central[0] + 14)
  assert.match(zipDiff(base, c), /central header fields/)
  const l = Buffer.from(base); l.writeUInt16LE(0x5a21, offsets(l).local[1] + 12)
  assert.match(zipDiff(base, l), /local header fields/)
})

test("zipDiff: a CRC32 that doesn't match the data is a difference", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  // both zips carry the same bogus CRC (header fields equal), so only the data check catches it
  const a = Buffer.from(base), b = Buffer.from(base)
  for (const z of [a, b]) { const o = offsets(z); for (const p of o.local) z.writeUInt32LE(0xdeadbeef, p + 14); for (const p of o.central) z.writeUInt32LE(0xdeadbeef, p + 16) }
  assert.match(zipDiff(a, b), /CRC32 doesn't match/)
  assert.match(zipDiff(base, a), /header fields|CRC32/)
})

test("zipDiff: an archive comment or trailing bytes are a difference", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const withComment = Buffer.concat([base, Buffer.from("hi")])
  withComment.writeUInt16LE(2, offsets(withComment).eocd + 20)
  assert.match(zipDiff(base, withComment), /archive comment/)
  const trailing = Buffer.concat([base, Buffer.from("junk")])
  assert.match(zipDiff(base, trailing), /trailing bytes/)
})

test("compareDirs reports tampered zips as DIFF with the reason, recompressed ones as same", async () => {
  const { zip, compareDirs, parts, base, offsets } = await zipFixtures()
  const g = fs.mkdtempSync(path.join(os.tmpdir(), "zipdiff-g-")), c = fs.mkdtempSync(path.join(os.tmpdir(), "zipdiff-c-"))
  try {
    const t = Buffer.from(base); t.writeUInt16LE(0x5a21, offsets(t).central[0] + 14)
    fs.writeFileSync(path.join(g, "re.xlsx"), base); fs.writeFileSync(path.join(c, "re.xlsx"), zip(parts, { level: 1 }))
    fs.writeFileSync(path.join(g, "ts.xlsx"), base); fs.writeFileSync(path.join(c, "ts.xlsx"), t)
    const r = compareDirs(g, c)
    assert.deepEqual(r.zipRecompressed, ["re.xlsx"])
    assert.deepEqual(r.differ, ["ts.xlsx"])
    assert.match(r.reasons["ts.xlsx"], /central header fields/)
  } finally { fs.rmSync(g, { recursive: true, force: true }); fs.rmSync(c, { recursive: true, force: true }) }
})

test("zipDiff: a local header whose method or compressed size disagrees with its central header", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const m = Buffer.from(base); m.writeUInt16LE(0, offsets(m).local[0] + 8) // local says stored, central says deflate
  assert.match(zipDiff(base, m), /not a plain zip/)
  const c = Buffer.from(base); c.writeUInt32LE(c.readUInt32LE(offsets(c).local[0] + 18) + 1, offsets(c).local[0] + 18)
  assert.match(zipDiff(base, c), /not a plain zip/)
})

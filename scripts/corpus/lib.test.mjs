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

test("zipDiff: junk at the end of the central directory", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const o = offsets(base)
  // 4 junk bytes between the last central header and the EOCD; the EOCD's size grows to cover them
  const j = Buffer.concat([base.subarray(0, o.eocd), Buffer.from("JUNK"), base.subarray(o.eocd)])
  j.writeUInt32LE(j.readUInt32LE(o.eocd + 4 + 12) + 4, o.eocd + 4 + 12)
  assert.match(zipDiff(base, j), /not a plain zip/)
})

test("zipDiff: EOCD disk number changed", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const d = Buffer.from(base); d.writeUInt16LE(1, offsets(d).eocd + 4) // number of this disk
  assert.match(zipDiff(base, d), /end-of-central-directory fields/)
})

test("zipDiff: junk between the entries and the central directory (offset moved)", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const o = offsets(base), cd = base.readUInt32LE(o.eocd + 16)
  const j = Buffer.concat([base.subarray(0, cd), Buffer.from("JUNK"), base.subarray(cd)])
  j.writeUInt32LE(cd + 4, o.eocd + 4 + 16) // central directory offset, now 4 bytes later
  assert.match(zipDiff(base, j), /not a plain zip/)
})

test("zipDiff: archive comment length past the end of the file", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const c = Buffer.from(base); c.writeUInt16LE(10, offsets(c).eocd + 20) // claims 10 bytes; none follow
  assert.match(zipDiff(base, c), /archive comment runs past the end/)
})

test("zipDiff: an uncompressed size that doesn't match the data", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  // the same wrong size in both zips' headers (fields equal), so only the data check catches it
  const a = Buffer.from(base), b = Buffer.from(base)
  for (const z of [a, b]) { const o = offsets(z); z.writeUInt32LE(5 + 1, o.local[1] + 22); z.writeUInt32LE(5 + 1, o.central[1] + 24) } // b.txt is 5 bytes
  assert.match(zipDiff(a, b), /uncompressed size doesn't match/)
})

test("zipDiff: junk between two entries (later offsets shifted)", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const o = offsets(base), gap = o.local[1] // entry 1 starts right where entry 0 ends
  const j = Buffer.concat([base.subarray(0, gap), Buffer.from("JUNK"), base.subarray(gap)])
  // everything from entry 1 on moved by 4 bytes; fix the offsets that point past the gap
  j.writeUInt32LE(j.readUInt32LE(o.central[1] + 4 + 42) + 4, o.central[1] + 4 + 42) // entry 1's local header offset
  j.writeUInt32LE(j.readUInt32LE(o.eocd + 4 + 16) + 4, o.eocd + 4 + 16) // central directory offset
  assert.equal(zipDiff(j, j), "not a plain zip (zip64, data descriptor or bad layout)") // readable, but not back to back
  assert.match(zipDiff(base, j), /not a plain zip/)
})

test("zipDiff: only the local uncompressed size is wrong", async () => {
  const { zipDiff, base, offsets } = await zipFixtures()
  const a = Buffer.from(base), b = Buffer.from(base)
  for (const z of [a, b]) z.writeUInt32LE(5 + 1, offsets(z).local[1] + 22) // central still says 5 (b.txt)
  assert.match(zipDiff(a, b), /uncompressed size doesn't match/)
})

test("childErrorDetail: the error line, not the location line Node prints first", async () => {
  const { childErrorDetail } = await import("./lib.mjs")
  const { execFileSync } = await import("node:child_process")
  let err
  try { execFileSync(process.execPath, ["--input-type=module", "-e", "new Date(NaN).toISOString()"], { stdio: ["ignore", "pipe", "pipe"] }) } catch (e) { err = e }
  assert.ok(err, "child should fail")
  assert.match(err.stderr.toString().trim().split("\n")[0], /^file:|\[eval/, "Node's first stderr line is a location") // what the old code reported
  assert.match(childErrorDetail(err), /^RangeError: Invalid time value/)
  // no error line: the last line (e.g. a message the child printed itself before exit 3)
  assert.equal(childErrorDetail({ stderr: Buffer.from("\nG2 child: TZ=Bogus/Zone gives a UTC offset of 0\n"), message: "Command failed" }), "G2 child: TZ=Bogus/Zone gives a UTC offset of 0")
  // no stderr at all: the error's message
  assert.equal(childErrorDetail({ stderr: Buffer.from(""), message: "spawnSync node ETIMEDOUT" }), "spawnSync node ETIMEDOUT")
})

test("childErrorDetail: a timeout or signal is kept when the child wrote to stderr", async () => {
  const { childErrorDetail } = await import("./lib.mjs")
  const { execFileSync } = await import("node:child_process")
  const fail = (args, opts = {}) => { try { execFileSync(process.execPath, args, { stdio: ["ignore", "pipe", "pipe"], ...opts }) } catch (e) { return e } assert.fail("child should fail") }
  // writes a line, then hangs until the timeout kills it
  const t = fail(["-e", 'process.stderr.write("partial output\\n"); setTimeout(() => {}, 60_000)'], { timeout: 500, killSignal: "SIGKILL" })
  assert.equal(childErrorDetail(t), "ETIMEDOUT: partial output")
  // killed by a signal
  const s = fail(["-e", 'process.stderr.write("about to die\\n"); process.kill(process.pid, "SIGTERM")'])
  assert.equal(childErrorDetail(s), "SIGTERM: about to die")
  // a plain non-zero exit has neither: no prefix
  const x = fail(["-e", 'process.stderr.write("bye\\n"); process.exit(3)'])
  assert.equal(childErrorDetail(x), "bye")
})

test("childErrorDetail: thrown strings, objects, DOMException and subclassed errors", async () => {
  const { childErrorDetail } = await import("./lib.mjs")
  const { execFileSync } = await import("node:child_process")
  const detail = (code) => { try { execFileSync(process.execPath, ["--input-type=module", "-e", code], { stdio: ["ignore", "pipe", "pipe"] }) } catch (e) { return childErrorDetail(e) } assert.fail("child should fail") }
  assert.equal(detail('throw "a thrown string"'), "a thrown string")
  assert.equal(detail('throw { custom: 1 }'), "{ custom: 1 }")
  assert.equal(detail('throw new DOMException("stop", "AbortError")'), "DOMException [AbortError]: stop")
  assert.equal(detail('class MyErr extends Error {}; throw new MyErr("mine")'), "MyErr [Error]: mine")
  // a source line that starts like an error name isn't taken for the error line
  assert.equal(detail('Error.captureStackTrace({}); throw new TypeError("bad type")'), "TypeError: bad type")
})

test("childErrorDetail: the last error line wins over earlier lines the child printed", async () => {
  const { childErrorDetail } = await import("./lib.mjs")
  const { execFileSync } = await import("node:child_process")
  const detail = (code) => { try { execFileSync(process.execPath, ["--input-type=module", "-e", code], { stdio: ["ignore", "pipe", "pipe"] }) } catch (e) { return childErrorDetail(e) } assert.fail("child should fail") }
  assert.equal(detail('console.error("debug: x"); throw new TypeError("bad")'), "TypeError: bad")
  assert.equal(detail('console.error("Warning: an earlier Error: not this one"); throw new RangeError("this one")'), "RangeError: this one")
  // Node prints error properties after the stack ("code: 'ERR_X'"); they aren't the error line
  assert.equal(detail('console.error("debug: x"); const e = new Error("boom"); e.code = "ERR_X"; throw e'), "Error: boom")
  // a child that printed an earlier error line and then exits on its own: the last one
  assert.equal(childErrorDetail({ stderr: Buffer.from("SyntaxError: first\ndebug: x\nTypeError: second\n"), message: "Command failed" }), "TypeError: second")
  // nothing error-like: the last non-empty line, then the message
  assert.equal(childErrorDetail({ stderr: Buffer.from("debug: x\nlast words\n\n"), message: "m" }), "last words")
  assert.equal(childErrorDetail({ stderr: Buffer.from(""), message: "m" }), "m")
})

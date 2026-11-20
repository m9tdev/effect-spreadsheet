/**
 * Differential check of the TypeScript inflate (`src/internal/zip/rfc1951.ts`)
 * against the platform `DecompressionStream` (the library's default) and
 * node:zlib, over every deflate entry of every zip in the corpus: the
 * committed fixtures, the gitignored `test/corpus/.cache` and the SheetJS
 * test files in `$SHEETJS_DIR` when present. Each entry is also checked cut
 * in half, cut by its last byte, with 4 bytes appended and with one bit
 * flipped. Not a test (the corpus cache is not always there).
 * Run: node scripts/bench/inflate-differential.ts [--list]
 *
 * For each case:
 * - the TypeScript inflate and `DecompressionStream` must agree on success
 *   or failure, and on the bytes when both succeed;
 * - when both fail, the platform's bytes before the error must be a prefix
 *   of the TypeScript inflate's (it may hand out more of them: it delivers
 *   every byte decoded before the error);
 * - against zlib (which ignores trailing data, and with a sync flush gives
 *   the bytes of a truncated stream without failing): the same bytes, and
 *   a failure exactly where zlib stops early;
 * - the TypeScript inflate gives the same bytes and the same end whether
 *   it is driven with 64 KiB steps or with 4099-byte outputs and 37-byte
 *   input budgets.
 */
import { readdirSync, readFileSync, statSync } from "node:fs"
import { join } from "node:path"
import * as zlib from "node:zlib"
import * as Result from "effect/Result"
import { defaults } from "../../src/SpreadsheetLimits.ts"
import { locateEntry } from "../../src/internal/zip/archive.ts"
import { readDirectory } from "../../src/internal/zip/directory.ts"
import { inflateRawWeb } from "../../src/internal/zip/inflate.ts"
import { inflate } from "../../src/internal/zip/rfc1951.ts"

const root = new URL("../../", import.meta.url).pathname
const dirs = [join(root, "test/corpus"), join(root, "test/fixtures")]
const sheetjs = process.env["SHEETJS_DIR"] ?? join(process.env["HOME"] ?? "", "sheetjs")
dirs.push(join(sheetjs, "test_files"))
const list = process.argv.includes("--list")

const files = (dir: string): Array<string> => {
  let entries
  try {
    entries = readdirSync(dir, { withFileTypes: true })
  } catch {
    return []
  }
  return entries.flatMap((e) => {
    const path = join(dir, e.name)
    if (e.isSymbolicLink()) {
      try {
        return statSync(path).isDirectory() ? files(path) : [path]
      } catch {
        return []
      }
    }
    return e.isDirectory() ? files(path) : e.isFile() ? [path] : []
  })
}

/** Output past this is not compared (some corpus entries are deflate bombs); such cases are counted as capped. */
const CAP = 256 * 1024 * 1024

const concat = (parts: ReadonlyArray<Uint8Array>) => Buffer.concat(parts.map((p) => Buffer.from(p.buffer, p.byteOffset, p.byteLength)))

interface Outcome {
  readonly ok: boolean
  readonly capped?: boolean
  readonly bytes: Buffer
  readonly reason?: string
}

const platform = async (data: Uint8Array): Promise<Outcome> => {
  const inflater = inflateRawWeb(data)
  const parts: Array<Uint8Array> = []
  let size = 0
  try {
    // progress: each call returns a chunk or ends
    for (;;) {
      const chunk = await inflater.next()
      if (chunk === undefined) return { ok: true, bytes: concat(parts) }
      parts.push(chunk)
      size += chunk.length
      if (size >= CAP) return { ok: true, capped: true, bytes: concat(parts).subarray(0, CAP) }
    }
  } catch (error) {
    return { ok: false, bytes: concat(parts), reason: String(error instanceof Error ? error.message : error) }
  } finally {
    inflater.cancel()
  }
}

const ours = (data: Uint8Array, outSize: number, maxInput: number): Outcome => {
  const decoder = inflate(data)
  const parts: Array<Uint8Array> = []
  let size = 0
  // progress: each step writes output, consumes input, or ends
  for (;;) {
    const out = new Uint8Array(outSize)
    const step = decoder.next(out, maxInput)
    if (step._tag === "Output") {
      parts.push(out.subarray(0, step.written))
      size += step.written
      if (size >= CAP) return { ok: true, capped: true, bytes: concat(parts).subarray(0, CAP) }
    }
    else if (step._tag === "Done") return { ok: true, bytes: concat(parts) }
    else if (step._tag === "Failed") return { ok: false, bytes: concat(parts), reason: `${step.reason} @${step.offset}` }
  }
}

const node = (data: Uint8Array): Outcome => {
  try {
    return { ok: true, bytes: zlib.inflateRawSync(data, { finishFlush: zlib.constants.Z_SYNC_FLUSH, maxOutputLength: CAP }) }
  } catch (error) {
    if (error instanceof RangeError) return { ok: true, capped: true, bytes: Buffer.alloc(0) }
    return { ok: false, bytes: Buffer.alloc(0), reason: String(error instanceof Error ? error.message : error) }
  }
}

const isPrefix = (a: Buffer, b: Buffer) => a.length <= b.length && b.subarray(0, a.length).equals(a)

const stats = {
  files: 0,
  zips: 0,
  entries: 0,
  cases: 0,
  bothOk: 0,
  bothFailed: 0,
  oursMoreBeforeError: 0,
  zlibTruncatedExact: 0,
  zlibTrailingExact: 0,
  compressedBytes: 0,
  capped: 0,
  diffs: 0
}
const diffs: Array<string> = []
const reasons = new Map<string, number>()

const check = async (label: string, data: Uint8Array) => {
  stats.cases++
  const p = await platform(data)
  const t = ours(data, 65536, 65536)
  const t2 = ours(data, 4099, 37)
  const z = node(data)
  const problems: Array<string> = []
  if (p.capped || t.capped || t2.capped || z.capped) {
    // only the first CAP bytes are compared
    stats.capped++
    const n = Math.min(p.bytes.length, t.bytes.length, t2.bytes.length)
    if (!p.bytes.subarray(0, n).equals(t.bytes.subarray(0, n)) || !t2.bytes.subarray(0, n).equals(t.bytes.subarray(0, n))) {
      stats.diffs++
      diffs.push(`${label}: different bytes in the first ${n} (capped)`)
    }
    return
  }
  if (t.ok !== t2.ok || !t.bytes.equals(t2.bytes) || t.reason !== t2.reason) problems.push("step sizes change the result")
  if (p.ok !== t.ok) problems.push(`platform ${p.ok ? "ok" : `failed (${p.reason})`}, ours ${t.ok ? "ok" : `failed (${t.reason})`}`)
  else if (p.ok) {
    stats.bothOk++
    if (!p.bytes.equals(t.bytes)) problems.push("different bytes")
  } else {
    stats.bothFailed++
    const key = (t.reason ?? "").replace(/ @\d+$/, "")
    reasons.set(key, (reasons.get(key) ?? 0) + 1)
    if (!isPrefix(p.bytes, t.bytes)) problems.push(`platform bytes before the error (${p.bytes.length}) not a prefix of ours (${t.bytes.length})`)
    else if (t.bytes.length > p.bytes.length) stats.oursMoreBeforeError++
  }
  if (z.ok) {
    if (t.ok) {
      if (!z.bytes.equals(t.bytes)) problems.push("zlib: different bytes")
    } else if (!z.bytes.equals(t.bytes)) problems.push(`zlib ok with ${z.bytes.length} bytes, ours failed (${t.reason}) after ${t.bytes.length}`)
    else if (t.reason?.startsWith("unexpected end")) stats.zlibTruncatedExact++
    else if (t.reason?.startsWith("data after the end")) stats.zlibTrailingExact++
    else problems.push(`zlib ok, ours failed (${t.reason}) after the same bytes`)
  } else if (t.ok) problems.push(`zlib failed (${z.reason}), ours ok`)
  if (problems.length > 0) {
    stats.diffs++
    diffs.push(`${label}: ${problems.join("; ")}`)
  }
}

let seed = 0x9E3779B9
const random = () => {
  seed = (seed * 1103515245 + 12345) >>> 0
  return seed
}

const started = performance.now()
for (const dir of dirs) {
  for (const path of files(dir)) {
    stats.files++
    if (stats.files % 1000 === 0) process.stderr.write(`${stats.files} files, ${stats.cases} cases, ${stats.diffs} diffs\n`)
    const bytes = new Uint8Array(readFileSync(path))
    const directory = readDirectory(bytes, { ...defaults, maxZipEntries: 1e6 })
    if (Result.isFailure(directory)) continue
    stats.zips++
    for (const entry of directory.success.entries) {
      if (entry.method !== 8) continue
      const located = locateEntry(bytes, directory.success, entry)
      if (Result.isFailure(located)) continue
      const data = bytes.subarray(located.success.data.start, located.success.data.end)
      stats.entries++
      stats.compressedBytes += data.length
      const label = `${path.slice(root.length)}!${entry.name}`
      await check(label, data)
      if (data.length > 0) {
        await check(`${label} (cut in half)`, data.subarray(0, data.length >>> 1))
        await check(`${label} (last byte cut)`, data.subarray(0, data.length - 1))
        const flipped = data.slice()
        const bit = random() % (data.length * 8)
        flipped[bit >>> 3] ^= 1 << (bit & 7)
        await check(`${label} (bit ${bit} flipped)`, flipped)
      }
      await check(`${label} (4 bytes appended)`, Buffer.concat([data, Buffer.from([1, 2, 3, 4])]))
    }
  }
}
console.log(JSON.stringify({ ...stats, seconds: Math.round((performance.now() - started) / 1000) }))
console.log("failure reasons (ours, both failed):", JSON.stringify(Object.fromEntries(reasons)))
for (const line of list ? diffs : diffs.slice(0, 50)) console.log("DIFF", line)

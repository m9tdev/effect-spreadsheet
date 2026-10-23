#!/usr/bin/env node
// Fetch the SheetJS test_files corpus into test/corpus/.cache/sheetjs (gitignored).
//
// Upstream: https://test-files.sheetjs.com/test_files.zip (the URL SheetJS's own
// `make init` uses). The GitHub repo SheetJS/test_files is blocked by GitHub and
// git.sheetjs.com has no public test_files repo, so there is no commit to pin:
// we pin the archive by sha256 instead. If upstream changes the archive the
// script stops; re-run with --update to accept the new archive, which rewrites
// the pin in test/corpus/sources.json (commit it, and update docs/corpus.md).
//
// Idempotent: does nothing when the stamp file matches the pinned archive.
// A download that doesn't match the pin is deleted (so the next run downloads again).
// Usage: node scripts/corpus/fetch.mjs [--update] [--keep-zip] [--timeout SEC] (download timeout, default 900)

import fs from "node:fs"
import path from "node:path"
import { Readable } from "node:stream"
import { pipeline } from "node:stream/promises"
import { CACHE_DIR, SHEETJS_CACHE, SOURCES_PATH, sha256, zipEntries, zipRead } from "./lib.mjs"

const sources = JSON.parse(fs.readFileSync(SOURCES_PATH, "utf8"))
const pin = sources.sheetjs
const URL = pin.url
const PINNED_SHA256 = pin.archiveSha256
const PINNED_SIZE = pin.archiveSize

const argv = process.argv.slice(2)
const args = new Set(argv)
const ti = argv.indexOf("--timeout")
const DOWNLOAD_TIMEOUT_MS = ti >= 0 ? Number(argv[ti + 1]) * 1000 : 900_000
if (!(DOWNLOAD_TIMEOUT_MS > 0)) { console.error("--timeout needs a positive number of seconds"); process.exit(2) }
for (const [i, a] of argv.entries()) {
  if (!["--update", "--keep-zip", "--timeout"].includes(a) && argv[i - 1] !== "--timeout") { console.error(`unknown argument ${a}`); process.exit(2) }
}
const zipPath = path.join(CACHE_DIR, "test_files.zip")
const stampPath = path.join(SHEETJS_CACHE, ".fetched.json")

const t0 = Date.now()
fs.mkdirSync(CACHE_DIR, { recursive: true })

if (fs.existsSync(stampPath)) {
  const stamp = JSON.parse(fs.readFileSync(stampPath, "utf8"))
  if (stamp.sha256 === PINNED_SHA256 && !args.has("--update")) {
    console.log(`sheetjs corpus already present (${stamp.files} files, sha256 ${stamp.sha256.slice(0, 12)}…)`)
    process.exit(0)
  }
}

let lastModified = null
if (!fs.existsSync(zipPath) || fs.statSync(zipPath).size !== PINNED_SIZE || args.has("--update")) {
  console.log(`downloading ${URL} (timeout ${DOWNLOAD_TIMEOUT_MS / 1000}s) …`)
  const tmp = zipPath + ".part"
  try {
    const res = await fetch(URL, { signal: AbortSignal.timeout(DOWNLOAD_TIMEOUT_MS) })
    if (!res.ok || !res.body) throw new Error(`HTTP ${res.status}`)
    lastModified = res.headers.get("last-modified")
    await pipeline(Readable.fromWeb(res.body), fs.createWriteStream(tmp))
  } catch (e) {
    fs.rmSync(tmp, { force: true })
    console.error(`download failed: ${e?.name === "TimeoutError" ? `no complete download within ${DOWNLOAD_TIMEOUT_MS / 1000}s` : e?.message ?? e}`)
    process.exit(1)
  }
  fs.renameSync(tmp, zipPath)
}

const zip = fs.readFileSync(zipPath)
const digest = sha256(zip)
if (digest !== PINNED_SHA256 && !args.has("--update")) {
  fs.rmSync(zipPath, { force: true }) // don't keep a bad or unexpected archive around
  console.error(`sha256 mismatch: got ${digest} (${zip.length} bytes), pinned ${PINNED_SHA256} (${PINNED_SIZE} bytes); deleted the download.\n` +
    `Either the download was corrupt (just re-run) or upstream changed the archive (re-run with --update to accept it; that rewrites test/corpus/sources.json, commit it).`)
  process.exit(1)
}

// Extract into a fresh directory, then swap into place.
const staging = SHEETJS_CACHE + ".staging"
fs.rmSync(staging, { recursive: true, force: true })
fs.mkdirSync(staging, { recursive: true })
const entries = zipEntries(zip)
if (!entries) throw new Error("not a zip")
let files = 0, bytes = 0
for (const e of entries) {
  if (e.name.endsWith("/")) continue
  const norm = path.posix.normalize(e.name)
  if (norm.startsWith("../") || path.posix.isAbsolute(norm) || norm.startsWith("__MACOSX/")) continue
  const out = path.join(staging, norm)
  fs.mkdirSync(path.dirname(out), { recursive: true })
  const data = zipRead(zip, e, 2 ** 31 - 1)
  fs.writeFileSync(out, data)
  files++
  bytes += data.length
}
fs.writeFileSync(path.join(staging, ".fetched.json"), JSON.stringify({ url: URL, sha256: digest, size: zip.length, files, bytes, fetchedAt: new Date().toISOString() }, null, 2) + "\n")
fs.rmSync(SHEETJS_CACHE, { recursive: true, force: true })
fs.renameSync(staging, SHEETJS_CACHE)
if (!args.has("--keep-zip")) fs.rmSync(zipPath, { force: true })
console.log(`extracted ${files} files (${(bytes / 1e6).toFixed(0)} MB) into ${path.relative(process.cwd(), SHEETJS_CACHE)} in ${((Date.now() - t0) / 1000).toFixed(1)}s`)
if (digest !== PINNED_SHA256) {
  sources.sheetjs = { ...pin, archiveSha256: digest, archiveSize: zip.length, files, ...(lastModified ? { lastModified: new Date(lastModified).toISOString().replace(".000Z", "Z") } : {}) }
  fs.writeFileSync(SOURCES_PATH, JSON.stringify(sources, null, 2) + "\n")
  console.log(`NOTE: new archive sha256 ${digest}; updated ${path.relative(process.cwd(), SOURCES_PATH)} (commit it)`)
}

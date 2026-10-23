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
// Usage: node scripts/corpus/fetch.mjs [--update] [--keep-zip]

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

const args = new Set(process.argv.slice(2))
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

if (!fs.existsSync(zipPath) || fs.statSync(zipPath).size !== PINNED_SIZE || args.has("--update")) {
  console.log(`downloading ${URL} …`)
  const res = await fetch(URL)
  if (!res.ok || !res.body) throw new Error(`download failed: HTTP ${res.status}`)
  const tmp = zipPath + ".part"
  await pipeline(Readable.fromWeb(res.body), fs.createWriteStream(tmp))
  fs.renameSync(tmp, zipPath)
}

const zip = fs.readFileSync(zipPath)
const digest = sha256(zip)
if (digest !== PINNED_SHA256 && !args.has("--update")) {
  console.error(`sha256 mismatch: got ${digest}, pinned ${PINNED_SHA256}.\nUpstream archive changed; re-run with --update to accept it, then update PINNED_SHA256.`)
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
  sources.sheetjs = { ...pin, archiveSha256: digest, archiveSize: zip.length, files }
  fs.writeFileSync(SOURCES_PATH, JSON.stringify(sources, null, 2) + "\n")
  console.log(`NOTE: new archive sha256 ${digest}; updated ${path.relative(process.cwd(), SOURCES_PATH)} (commit it)`)
}

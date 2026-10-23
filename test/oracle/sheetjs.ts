/**
 * SheetJS oracle: loads the reference clone from ~/sheetjs/xlsx.mjs (never a
 * package dependency). `load()` gives undefined when it's absent, so tests
 * can skip.
 */
import { existsSync } from "node:fs"
import { homedir } from "node:os"
import { join } from "node:path"
import { pathToFileURL } from "node:url"

export interface SheetJS {
  read(data: Uint8Array, options: Record<string, unknown>): {
    SheetNames: Array<string>
    Sheets: Record<string, unknown>
  }
  utils: {
    sheet_to_json(sheet: unknown, options: Record<string, unknown>): Array<unknown>
  }
}

export const SHEETJS_PATH = process.env["SHEETJS_PATH"] ?? join(homedir(), "sheetjs", "xlsx.mjs")

let cached: Promise<SheetJS | undefined> | undefined

export const load = (): Promise<SheetJS | undefined> => {
  if (cached === undefined) {
    cached = existsSync(SHEETJS_PATH)
      ? import(pathToFileURL(SHEETJS_PATH).href).then((m: { default?: SheetJS } & SheetJS) => m.default ?? m)
      : Promise.resolve(undefined)
  }
  return cached
}

/** First sheet as `sheet_to_json(ws, { header: 1, raw: true })`, holes as undefined. */
export const aoa = (
  xlsx: SheetJS,
  bytes: Uint8Array,
  readOptions: Record<string, unknown> = {}
): Array<Array<unknown>> => {
  const wb = xlsx.read(bytes, { type: "buffer", ...readOptions })
  const ws = wb.Sheets[wb.SheetNames[0]!]
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: true }) as Array<Array<unknown>>
  return rows.map((row) => Array.from(row))
}

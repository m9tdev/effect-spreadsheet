import { readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"

const srcDir = new URL("../src", import.meta.url).pathname
const files = (dir: string): Array<string> =>
  readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
    e.isDirectory() ? files(join(dir, e.name)) : e.name.endsWith(".ts") ? [join(dir, e.name)] : []
  )

describe("src rules (AGENTS.md)", () => {
  // a package import of SheetJS / "xlsx" (not a relative path, not our own subpath)
  const PACKAGE_XLSX = /from\s+["'](?![./]|effect-spreadsheet\/)[^"']*(xlsx|sheetjs)[^"']*["']/i
  // anything named sheetjs, relative paths included
  const ANY_SHEETJS = /from\s+["'][^"']*sheetjs/i

  it("the import patterns catch what they should, and allow our own modules", () => {
    expect('import X from "xlsx"').toMatch(PACKAGE_XLSX)
    expect('import X from "../../sheetjs/xlsx.mjs"').toMatch(ANY_SHEETJS)
    expect('import X from "./xlsx/index.ts"').not.toMatch(PACKAGE_XLSX)
    expect('import X from "effect-spreadsheet/xlsx"').not.toMatch(PACKAGE_XLSX)
  })

  it("no node: imports, no require, no SheetJS import under src", () => {
    for (const file of files(srcDir)) {
      const text = readFileSync(file, "utf8")
      expect(text, file).not.toMatch(/from\s+["']node:|import\(\s*["']node:|\brequire\(/)
      expect(text, file).not.toMatch(PACKAGE_XLSX)
      expect(text, file).not.toMatch(ANY_SHEETJS)
    }
  })

  // Effect's own yield however it is reached: a namespace (`Effect.`, an alias),
  // yieldNowWith, or a named import from effect
  const EFFECT_YIELD = /\.\s*yieldNow|\byieldNowWith\b|import\s*\{[^}]*\byieldNow\b[^}]*\}\s*from\s*["']effect/

  it("the Effect yield pattern catches its forms, and allows the readers' own yield", () => {
    expect("yield* Effect.yieldNow").toMatch(EFFECT_YIELD)
    expect("yield* E.yieldNow").toMatch(EFFECT_YIELD)
    expect("Effect.yieldNowWith(1)").toMatch(EFFECT_YIELD)
    expect('import { Stream, yieldNow } from "effect/Effect"').toMatch(EFFECT_YIELD)
    expect("yield* yieldNow").not.toMatch(EFFECT_YIELD)
    expect('import { yieldNow } from "../yield.ts"').not.toMatch(EFFECT_YIELD)
  })

  it("readers yield through internal/yield.ts, not Effect.yieldNow (clamped to 4 ms in browsers)", () => {
    for (const file of files(srcDir)) {
      if (file.endsWith(join("internal", "yield.ts"))) continue
      expect(readFileSync(file, "utf8"), file).not.toMatch(EFFECT_YIELD)
    }
  })

  it("no third-party licence or derived-code markers under src (plain MIT)", () => {
    for (const file of files(srcDir)) {
      const text = readFileSync(file, "utf8")
      expect(text, file).not.toMatch(/derived from SheetJS/i)
      expect(text, file).not.toMatch(/Portions derived/i)
      expect(text, file).not.toMatch(/Apache/i)
    }
  })
})

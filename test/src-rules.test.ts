import { readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"

const srcDir = new URL("../src", import.meta.url).pathname
const files = (dir: string): Array<string> =>
  readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
    e.isDirectory() ? files(join(dir, e.name)) : e.name.endsWith(".ts") ? [join(dir, e.name)] : []
  )

describe("src rules (AGENTS.md)", () => {
  it("no node: imports, no require, no SheetJS import under src", () => {
    for (const file of files(srcDir)) {
      const text = readFileSync(file, "utf8")
      expect(text, file).not.toMatch(/from\s+["']node:|import\(\s*["']node:|\brequire\(/)
      expect(text, file).not.toMatch(/from\s+["'][^"']*(xlsx|sheetjs)[^"']*["']/i)
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

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

  it("files mentioning SheetJS derivation start with the header and are in NOTICE", () => {
    const notice = readFileSync(new URL("../NOTICE", import.meta.url), "utf8")
    for (const file of files(srcDir)) {
      const text = readFileSync(file, "utf8")
      if (!text.startsWith("// Portions derived from SheetJS")) continue
      expect(text.split("\n")[0], file).toMatch(/^\/\/ Portions derived from SheetJS \(bits\/[^)]+\), Copyright \(C\) SheetJS LLC, Apache-2\.0\.$/)
      expect(text.split("\n")[1], file).toBe("// Licensed under the Apache License, Version 2.0: see LICENSE-APACHE-2.0.")
      expect(notice, file).toContain(file.slice(file.indexOf("/src/") + 1))
    }
  })
})

import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

// Rule: effect and any @effect/* package used at runtime are peerDependencies
// ("^4.0.6") and pinned devDependencies ("4.0.6"), never dependencies (a
// second copy of effect breaks its runtime type checks).
const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
  license?: string
  files?: Array<string>
  dependencies?: Record<string, string>
  peerDependencies?: Record<string, string>
  devDependencies?: Record<string, string>
}

const isEffect = (name: string) => name === "effect" || name.startsWith("@effect/")

describe("package.json dependencies", () => {
  it("has no effect or @effect/* in dependencies", () => {
    const offending = Object.keys(pkg.dependencies ?? {}).filter(isEffect)
    expect(offending).toEqual([])
  })

  it("has effect as a ^4.0.6 peer dependency", () => {
    expect(pkg.peerDependencies?.["effect"]).toBe("^4.0.6")
  })

  it("pins effect packages in devDependencies to 4.0.6", () => {
    const dev = Object.entries(pkg.devDependencies ?? {}).filter(([name]) => isEffect(name))
    expect(dev.length).toBeGreaterThan(0)
    for (const [name, range] of dev) expect([name, range]).toEqual([name, "4.0.6"])
  })

  it("lists every peer effect package in devDependencies too", () => {
    for (const name of Object.keys(pkg.peerDependencies ?? {}).filter(isEffect)) {
      expect(pkg.devDependencies?.[name], name).toBe("4.0.6")
    }
  })
})

// Rule: the package is MIT, and the "files" allowlist ships only the library
// (build output, sources, docs and licence files), never tests, scripts or
// the test corpus.
describe("package.json licence and files", () => {
  it("is MIT", () => {
    expect(pkg.license).toBe("MIT")
  })

  it("ships only the allowlisted paths", () => {
    expect(Array.isArray(pkg.files)).toBe(true)
    const allowed = new Set(["dist", "src", "README.md", "LICENSE", "NOTICE"])
    for (const entry of pkg.files ?? []) expect(allowed.has(entry.replace(/^\.\/|\/$/g, "")), entry).toBe(true)
  })

  it("never ships test/, scripts/ or docs/", () => {
    for (const entry of pkg.files ?? []) {
      expect(entry, entry).not.toMatch(/^(\.\/)?(test|scripts|docs)(\/|$)/)
      expect(entry, entry).not.toMatch(/[*?]|^\.?\/?$/)
    }
  })
})

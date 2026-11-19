/**
 * Preload for the benches: makes the zip archive use the TypeScript inflate
 * (`src/internal/zip/rfc1951.ts`) instead of `DecompressionStream`, without
 * changing the library's default. Only the archive's import of
 * `./inflate.ts` is redirected, to a module that exports the TypeScript
 * inflater under the default's name.
 * Run: node --import ./scripts/bench/own-inflate.ts scripts/bench/xlsx-rows.ts
 */
import { registerHooks } from "node:module"

const archive = new URL("../../src/internal/zip/archive.ts", import.meta.url).href
const swapped = new URL("./own-inflate-default.ts", import.meta.url).href

registerHooks({
  resolve: (specifier, context, nextResolve) => {
    if (context.parentURL !== archive || specifier !== "./inflate.ts") return nextResolve(specifier, context)
    process.stderr.write("inflate: TypeScript (rfc1951.ts)\n")
    return { url: swapped, shortCircuit: true }
  }
})

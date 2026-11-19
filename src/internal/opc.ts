/**
 * The package (OPC) layer shared by XLSX and XLSB: parts read as XML, and
 * the relationships that link them.
 *
 * @internal
 */
import { Effect } from "effect"
import type { SpreadsheetError } from "../SpreadsheetError.ts"
import type { Limits } from "../SpreadsheetLimits.ts"
import { parse } from "./xml/parse.ts"
import { attr, type XmlHandler } from "./xml/tokenizer.ts"
import type { Archive, LocatedEntry, ReaderKey } from "./zip/archive.ts"
import { normalizeName } from "./zip/directory.ts"

/**
 * Who reads a package's parts, for the archive's inflate budget (see
 * `ReaderKey`): the reads made while the workbook opens (its part,
 * relationships, shared strings, styles), each sheet by its index, and each
 * external book by its number. A part two of them read is charged to each.
 */
export const readers = {
  open: "open",
  sheet: (index: number): ReaderKey => `sheet:${index}`,
  book: (book: number): ReaderKey => `book:${book}`
} as const

/** Tokenizes one part with `handler`, as one of the reads made while the workbook opens. */
export const parsePart = (
  archive: Archive,
  entry: LocatedEntry,
  handler: XmlHandler,
  limits: Limits
): Effect.Effect<void, SpreadsheetError> => parse(archive.read(entry, readers.open), handler, limits, entry.name)

/**
 * What a relationship points at.
 * - `Part`: a part of this package (the entry is its identity);
 * - `Missing`: a part name with no entry in the package;
 * - `External`: `TargetMode="External"`, outside the package;
 * - `Invalid`: a target that names no part (it climbs above the package
 *   root, has a bad `%`-escape, or names the root itself).
 */
export type Target =
  | { readonly _tag: "Part"; readonly entry: LocatedEntry }
  | { readonly _tag: "Missing"; readonly name: string }
  | { readonly _tag: "External" }
  | { readonly _tag: "Invalid"; readonly reason: string }

/** A target that names a part of the package, present or not. */
export type Named = Extract<Target, { readonly _tag: "Part" | "Missing" }>

const isNamed = (target: Target): target is Named => target._tag === "Part" || target._tag === "Missing"

/** The part named `name` in `archive`: `Part` if it has the entry, else `Missing`. */
export const lookup = (archive: Archive, name: string): Named => {
  const entry = archive.find(name)
  return entry === undefined ? { _tag: "Missing", name } : { _tag: "Part", entry }
}

/**
 * Resolves a relationship target against the directory of its source part
 * (`base`, its `/`-separated segments) by OPC part-name rules: `.`/`..`
 * segments, percent-encoding, a leading `/` from the package root. The part
 * name, or why the target names none.
 */
export const resolveName = (
  base: ReadonlyArray<string>,
  target: string
): string | { readonly _tag: "Invalid"; readonly reason: string } => {
  const out = target.startsWith("/") ? [] : [...base]
  // progress: one segment per iteration
  for (const raw of target.split("/")) {
    let segment: string
    try {
      segment = decodeURIComponent(raw)
    } catch {
      return { _tag: "Invalid", reason: `${JSON.stringify(target)} has a bad %-escape` }
    }
    if (segment === "" || segment === ".") continue
    if (segment === "..") {
      if (out.length === 0) return { _tag: "Invalid", reason: `${JSON.stringify(target)} climbs above the package root` }
      out.pop()
      continue
    }
    out.push(segment)
  }
  return out.length === 0 ? { _tag: "Invalid", reason: `${JSON.stringify(target)} names no part` } : out.join("/")
}

/** The relationships part of a part: `xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`. */
export const relsPartOf = (part: string): string => {
  const slash = part.lastIndexOf("/")
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`
}

export interface Relationship {
  readonly id: string
  /**
   * The last segment of the type URI (`.../officeDocument` ->
   * `officeDocument`), so transitional and strict OOXML both match.
   */
  readonly kind: string
  readonly target: Target
}

/** The relationships of one source (the package, or a part). */
export interface Rels {
  /** The relationship with this id (the last one, if the id repeats). */
  readonly byId: (id: string) => Relationship | undefined
  /** The target of the first relationship of this kind that names a part, present or not. */
  readonly named: (kind: string) => Named | undefined
  /** The part of `named(kind)`; `undefined` if there's no such relationship or its part is missing. */
  readonly part: (kind: string) => LocatedEntry | undefined
}

/** The last path segment of a relationship type URI (`.../officeDocument` -> `officeDocument`). */
export const relationshipKind = (type: string): string => type.slice(type.lastIndexOf("/") + 1)

const NO_RELS: Rels = { byId: () => undefined, named: () => undefined, part: () => undefined }

/** Reads the relationships part `relsPart` (none if it's absent), resolving targets against `base`. */
const readRels = (
  archive: Archive,
  relsPart: string,
  base: ReadonlyArray<string>,
  limits: Limits
): Effect.Effect<Rels, SpreadsheetError> => {
  const entry = archive.find(relsPart)
  if (entry === undefined) return Effect.succeed(NO_RELS)
  const relationships: Array<Relationship> = []
  const handler: XmlHandler = {
    open: (name, attributes) => {
      if (name !== "Relationship") return
      const id = attr(attributes, "Id")
      const type = attr(attributes, "Type")
      const target = attr(attributes, "Target")
      if (id === undefined || type === undefined || target === undefined) return
      let resolved: Target
      if (attr(attributes, "TargetMode") === "External") {
        resolved = { _tag: "External" }
      } else {
        const name = resolveName(base, target)
        resolved = typeof name === "string" ? lookup(archive, name) : name
      }
      relationships.push({ id, kind: relationshipKind(type), target: resolved })
    },
    close: () => undefined,
    text: () => undefined
  }
  return Effect.map(parsePart(archive, entry, handler, limits), () => {
    const byId = new Map(relationships.map((r) => [r.id, r]))
    const named = (kind: string): Named | undefined => {
      // progress: one relationship per iteration
      for (const r of relationships) {
        if (r.kind === kind && isNamed(r.target)) return r.target
      }
      return undefined
    }
    return {
      byId: (id) => byId.get(id),
      named,
      part: (kind) => {
        const target = named(kind)
        return target?._tag === "Part" ? target.entry : undefined
      }
    }
  })
}

/** The package relationships (`_rels/.rels`), targets resolved from the package root. */
export const packageRels = (archive: Archive, limits: Limits): Effect.Effect<Rels, SpreadsheetError> =>
  readRels(archive, "_rels/.rels", [], limits)

/** The relationships of `part`, targets resolved against its directory. */
export const partRels = (archive: Archive, part: LocatedEntry, limits: Limits): Effect.Effect<Rels, SpreadsheetError> =>
  readRels(archive, relsPartOf(part.name), part.name.split("/").slice(0, -1), limits)

/**
 * The content type `[Content_Types].xml` gives `part`: the first `Override`
 * naming it, else the first `Default` for its extension; `undefined` if
 * there's neither, or no content types part. Part names match as the
 * archive matches entry names (`%`-escapes decoded, case-insensitive),
 * extensions case-insensitively.
 */
export const contentType = (
  archive: Archive,
  part: LocatedEntry,
  limits: Limits
): Effect.Effect<string | undefined, SpreadsheetError> => {
  const entry = archive.find("[Content_Types].xml")
  if (entry === undefined) return Effect.undefined
  const key = normalizeName(part.name)
  const dot = key.lastIndexOf(".")
  const extension = dot > key.lastIndexOf("/") ? key.slice(dot + 1) : undefined
  let override: string | undefined
  let byExtension: string | undefined
  const handler: XmlHandler = {
    open: (name, attributes) => {
      if (name === "Override" && override === undefined) {
        const partName = attr(attributes, "PartName")
        const type = attr(attributes, "ContentType")
        if (partName === undefined || type === undefined) return
        const resolved = resolveName([], partName)
        if (typeof resolved === "string" && normalizeName(resolved) === key) override = type
      } else if (name === "Default" && byExtension === undefined && extension !== undefined) {
        const type = attr(attributes, "ContentType")
        if (type !== undefined && attr(attributes, "Extension")?.toLowerCase() === extension) byExtension = type
      }
    },
    close: () => undefined,
    text: () => undefined
  }
  return Effect.map(parsePart(archive, entry, handler, limits), () => override ?? byExtension)
}

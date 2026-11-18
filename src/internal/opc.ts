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
import type { Archive } from "./zip/archive.ts"
import type { ZipEntry } from "./zip/directory.ts"

/** Tokenizes one part with `handler`. */
export const parsePart = (
  archive: Archive,
  entry: ZipEntry,
  handler: XmlHandler,
  limits: Limits
): Effect.Effect<void, SpreadsheetError> => parse(archive.read(entry), handler, limits, entry.name)

/**
 * Resolves a relationship target against the directory of its source part
 * (OPC part names: `/` separated, `.`/`..` segments, percent-encoding).
 * `undefined` if it climbs above the package root or isn't decodable.
 */
export const resolvePart = (sourcePart: string, target: string): string | undefined => {
  const base = target.startsWith("/") ? [] : sourcePart.split("/").slice(0, -1)
  const out = [...base]
  // progress: one segment per iteration
  for (const raw of target.split("/")) {
    let segment: string
    try {
      segment = decodeURIComponent(raw)
    } catch {
      return undefined
    }
    if (segment === "" || segment === ".") continue
    if (segment === "..") {
      if (out.length === 0) return undefined
      out.pop()
      continue
    }
    out.push(segment)
  }
  return out.length === 0 ? undefined : out.join("/")
}

/** The relationships part of a part: `xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`. */
export const relsPartOf = (part: string): string => {
  const slash = part.lastIndexOf("/")
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`
}

export interface Relationship {
  readonly id: string
  /** The type URI. Compared by its last segment, so transitional and strict OOXML both match. */
  readonly type: string
  /** The resolved part name; `undefined` for external or unresolvable targets. */
  readonly part: string | undefined
}

/** The last path segment of a relationship type URI (`.../officeDocument` -> `officeDocument`). */
export const relationshipKind = (type: string): string => type.slice(type.lastIndexOf("/") + 1)

/** Reads the relationships of `sourcePart` (an empty list if the part has none). */
export const readRelationships = (
  archive: Archive,
  sourcePart: string,
  limits: Limits
): Effect.Effect<ReadonlyArray<Relationship>, SpreadsheetError> => {
  const entry = archive.find(sourcePart === "" ? "_rels/.rels" : relsPartOf(sourcePart))
  if (entry === undefined) return Effect.succeed([])
  const relationships: Array<Relationship> = []
  const handler: XmlHandler = {
    open: (name, attributes) => {
      if (name !== "Relationship") return
      const id = attr(attributes, "Id")
      const type = attr(attributes, "Type")
      const target = attr(attributes, "Target")
      if (id === undefined || type === undefined || target === undefined) return
      const external = attr(attributes, "TargetMode") === "External"
      relationships.push({ id, type, part: external ? undefined : resolvePart(sourcePart, target) })
    },
    close: () => undefined,
    text: () => undefined
  }
  return Effect.as(parsePart(archive, entry, handler, limits), relationships)
}

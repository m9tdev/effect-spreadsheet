/**
 * The small XLSX parts read when a workbook is opened: relationships, the
 * workbook part, shared strings and styles. Each is tokenized as a stream
 * with the XML limits, so a huge or hostile part costs work in proportion
 * to its bytes and fails with a typed error.
 *
 * @internal
 */
import { Effect } from "effect"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { parse } from "../xml/parse.ts"
import type { XmlAttributes, XmlHandler } from "../xml/tokenizer.ts"
import type { Archive } from "../zip/archive.ts"
import type { ZipEntry } from "../zip/directory.ts"
import { formatFor, type NumberFormat } from "./numFmt.ts"

export const malformed = (reason: string, part?: string) =>
  new MalformedFile({ format: "xlsx", reason, ...(part === undefined ? {} : { part }) })

/** The value of attribute `name` (as written, e.g. `name`, `r:id`). */
export const attr = (attributes: XmlAttributes, name: string): string | undefined => {
  // progress: one attribute per iteration (bounded by maxXmlAttributes)
  for (const [key, value] of attributes) if (key === name) return value
  return undefined
}

/** A relationship id attribute: `id` with any namespace prefix (`r:id`, `x:id`). */
export const relationshipId = (attributes: XmlAttributes): string | undefined => {
  // progress: one attribute per iteration
  for (const [key, value] of attributes) {
    const colon = key.indexOf(":")
    if (colon > 0 && key.slice(colon + 1) === "id") return value
  }
  return undefined
}

/** `true`/`1` (xsd:boolean). */
export const isTrue = (value: string | undefined): boolean => value === "1" || value === "true"

/** Tokenizes one part with `handler`. */
export const parsePart = (
  archive: Archive,
  entry: ZipEntry,
  handler: XmlHandler,
  limits: Limits
): Effect.Effect<void, SpreadsheetError> => Effect.asVoid(parse(archive.read(entry), handler, limits, entry.name))

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

export type SheetKind = "worksheet" | "chartsheet" | "dialogsheet" | "macrosheet" | "other"

export interface SheetEntry {
  readonly name: string
  readonly visibility: "visible" | "hidden" | "veryHidden"
  readonly relationshipId: string | undefined
}

export interface DefinedNameEntry {
  readonly name: string
  readonly formula: string
  readonly scope: number | undefined
  readonly hidden: boolean
}

export interface WorkbookPart {
  readonly sheets: ReadonlyArray<SheetEntry>
  readonly definedNames: ReadonlyArray<DefinedNameEntry>
  readonly date1904: boolean
}

/** Reads the workbook part: sheets in order, defined names, the date system. */
export const readWorkbookPart = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<WorkbookPart, SpreadsheetError> => {
  const sheets: Array<SheetEntry> = []
  const definedNames: Array<DefinedNameEntry> = []
  let date1904 = false
  let inSheets = false
  let name: { name: string; scope: number | undefined; hidden: boolean; text: Array<string>; length: number } | undefined
  const handler: XmlHandler = {
    open: (element, attributes) => {
      switch (element) {
        case "workbookPr":
          date1904 = isTrue(attr(attributes, "date1904"))
          return
        case "sheets":
          inSheets = true
          return
        case "sheet": {
          if (!inSheets) return
          if (sheets.length >= limits.maxSheets) {
            return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: sheets.length + 1 })
          }
          const state = attr(attributes, "state")
          sheets.push({
            name: attr(attributes, "name") ?? `Sheet${sheets.length + 1}`,
            visibility: state === "hidden" ? "hidden" : state === "veryHidden" ? "veryHidden" : "visible",
            relationshipId: relationshipId(attributes)
          })
          return
        }
        case "definedName": {
          const local = attr(attributes, "localSheetId")
          const scope = local !== undefined && /^\d{1,9}$/.test(local) ? Number(local) : undefined
          name = { name: attr(attributes, "name") ?? "", scope, hidden: isTrue(attr(attributes, "hidden")), text: [], length: 0 }
          return
        }
      }
      return undefined
    },
    close: (element) => {
      if (element === "sheets") inSheets = false
      else if (element === "definedName" && name !== undefined) {
        definedNames.push({ name: name.name, formula: name.text.join(""), scope: name.scope, hidden: name.hidden })
        name = undefined
      }
    },
    text: (value) => {
      if (name === undefined) return
      name.text.push(value)
      name.length += value.length
      if (name.length > limits.maxStringLength) {
        return new LimitExceeded({ limit: "maxStringLength", max: limits.maxStringLength, actual: name.length })
      }
      return undefined
    }
  }
  return Effect.map(parsePart(archive, entry, handler, limits), () => ({ sheets, definedNames, date1904 }))
}

/** `_xHHHH_` escapes in OOXML strings (ST_Xstring): one UTF-16 code unit each; `_x005F_` is `_`. */
export const unescapeXstring = (s: string): string =>
  s.indexOf("_x") < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)))

/**
 * Collects rich text the way `<si>` and `<is>` hold it: the `<t>` elements
 * (directly, or inside `<r>` runs), joined; phonetic runs (`<rPh>`) are
 * skipped. Feed it the events inside the container.
 */
export class RichText {
  private parts: Array<string> = []
  private length = 0
  private inT = 0
  private inPhonetic = 0
  private readonly maxLength: number

  constructor(maxLength: number) {
    this.maxLength = maxLength
  }

  open(name: string): void {
    if (name === "rPh") this.inPhonetic++
    else if (name === "t" && this.inPhonetic === 0) this.inT++
  }

  close(name: string): void {
    if (name === "rPh") this.inPhonetic--
    else if (name === "t" && this.inT > 0) this.inT--
  }

  text(value: string): LimitExceeded | undefined {
    if (this.inT === 0) return undefined
    this.parts.push(value)
    this.length += value.length
    return this.length > this.maxLength
      ? new LimitExceeded({ limit: "maxStringLength", max: this.maxLength, actual: this.length })
      : undefined
  }

  /** The text so far (unescaped); resets for the next container. */
  take(): string {
    const value = unescapeXstring(this.parts.length === 1 ? this.parts[0]! : this.parts.join(""))
    this.parts = []
    this.length = 0
    this.inT = 0
    this.inPhonetic = 0
    return value
  }
}

/** Reads the shared-string table. */
export const readSharedStrings = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<ReadonlyArray<string>, SpreadsheetError> => {
  const strings: Array<string> = []
  const rich = new RichText(limits.maxStringLength)
  let inItem = false
  const handler: XmlHandler = {
    open: (name) => {
      if (name === "si") {
        if (strings.length >= limits.maxSharedStrings) {
          return new LimitExceeded({ limit: "maxSharedStrings", max: limits.maxSharedStrings, actual: strings.length + 1 })
        }
        inItem = true
        rich.take()
      } else if (inItem) rich.open(name)
      return undefined
    },
    close: (name) => {
      if (name === "si" && inItem) {
        strings.push(rich.take())
        inItem = false
      } else if (inItem) rich.close(name)
    },
    text: (value) => (inItem ? rich.text(value) : undefined)
  }
  return Effect.as(parsePart(archive, entry, handler, limits), strings)
}

/** Number formats by cell style (`<c s="...">` is an index into `cellXfs`). */
export interface Styles {
  readonly formats: ReadonlyArray<NumberFormat>
}

export const NO_STYLES: Styles = { formats: [] }

/** Reads the number formats and the cell styles that use them. */
export const readStyles = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<Styles, SpreadsheetError> => {
  const custom = new Map<number, string>()
  const xfIds: Array<number> = []
  let inCellXfs = false
  // custom codes only from <numFmts>: <dxfs><dxf><numFmt> (conditional
  // formats) may reuse an id with another code
  let inNumFmts = false
  const handler: XmlHandler = {
    open: (name, attributes) => {
      if (name === "numFmts") inNumFmts = true
      else if (name === "numFmt" && inNumFmts) {
        const id = attr(attributes, "numFmtId")
        const code = attr(attributes, "formatCode")
        if (id !== undefined && code !== undefined && /^\d{1,9}$/.test(id)) custom.set(Number(id), code)
      } else if (name === "cellXfs") inCellXfs = true
      else if (name === "xf" && inCellXfs) {
        const id = attr(attributes, "numFmtId")
        xfIds.push(id !== undefined && /^\d{1,9}$/.test(id) ? Number(id) : 0)
      }
    },
    close: (name) => {
      if (name === "cellXfs") inCellXfs = false
      else if (name === "numFmts") inNumFmts = false
    },
    text: () => undefined
  }
  return Effect.map(parsePart(archive, entry, handler, limits), () => ({
    formats: xfIds.map((id) => formatFor(id, custom))
  }))
}

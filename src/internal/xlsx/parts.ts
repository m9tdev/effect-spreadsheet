/**
 * The small XLSX parts read when a workbook is opened: the workbook part,
 * shared strings and styles. Each is tokenized as a stream
 * with the XML limits, so a huge or hostile part costs work in proportion
 * to its bytes and fails with a typed error.
 *
 * @internal
 */
import { Effect } from "effect"
import { LimitExceeded, type SpreadsheetError } from "../../SpreadsheetError.ts"
import type { Limits } from "../../SpreadsheetLimits.ts"
import { addFormat, cellStyles, type CellStyles, type NumberFormat } from "../numFmt.ts"
import { parsePart } from "../opc.ts"
import { malformedIn, type SheetEntry, sheetVisibility } from "../package.ts"
import { uniqueSheetNames } from "../sheetNames.ts"
import { attr, type XmlAttributes, type XmlHandler } from "../xml/tokenizer.ts"
import type { Archive } from "../zip/archive.ts"
import type { ZipEntry } from "../zip/directory.ts"

export const malformed = malformedIn("xlsx")

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

export type SheetKind = "worksheet" | "chartsheet" | "dialogsheet" | "macrosheet" | "other"

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

/**
 * Reads the workbook part: sheets in order, defined names, the date system.
 * Sheet names are decoded, then made unique (`uniqueSheetNames`).
 * A `<sheets>` inside `<sheets>`, or a `<definedName>` inside a defined
 * name, is MalformedFile: it would end the sheet list early, or replace the
 * open name.
 */
export const readWorkbookPart = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<WorkbookPart, SpreadsheetError> => {
  // names as the file gives them (decoded; `undefined` for a sheet without one), made unique at the end
  const sheets: Array<Omit<SheetEntry, "name"> & { readonly name: string | undefined }> = []
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
          if (inSheets) return malformed("<sheets> inside <sheets>", entry.name)
          inSheets = true
          return
        case "sheet": {
          if (!inSheets) return
          if (sheets.length >= limits.maxSheets) {
            return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: sheets.length + 1 })
          }
          const sheetName = attr(attributes, "name")
          sheets.push({
            name: sheetName === undefined ? undefined : unescapeXstring(sheetName),
            visibility: sheetVisibility(attr(attributes, "state")),
            relationshipId: relationshipId(attributes)
          })
          return
        }
        case "definedName": {
          if (name !== undefined) return malformed("<definedName> inside a <definedName>", entry.name)
          if (definedNames.length >= limits.maxDefinedNames) {
            return new LimitExceeded({ limit: "maxDefinedNames", max: limits.maxDefinedNames, actual: definedNames.length + 1 })
          }
          const local = attr(attributes, "localSheetId")
          const scope = local !== undefined && /^\d{1,9}$/.test(local) ? Number(local) : undefined
          name = { name: unescapeXstring(attr(attributes, "name") ?? ""), scope, hidden: isTrue(attr(attributes, "hidden")), text: [], length: 0 }
          return
        }
      }
      return undefined
    },
    close: (element) => {
      if (element === "sheets") inSheets = false
      else if (element === "definedName" && name !== undefined) {
        definedNames.push({ name: name.name, formula: unescapeXstring(name.text.join("")), scope: name.scope, hidden: name.hidden })
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
  return Effect.map(parsePart(archive, entry, handler, limits), () => {
    const names = uniqueSheetNames(sheets.map((s) => s.name))
    return { sheets: sheets.map((s, k) => ({ ...s, name: names[k]! })), definedNames, date1904 }
  })
}

/** `_xHHHH_` escapes in OOXML strings (ST_Xstring): one UTF-16 code unit each; `_x005F_` is `_`. */
export const unescapeXstring = (s: string): string =>
  s.indexOf("_x") < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)))

/**
 * Collects rich text the way `<si>` and `<is>` hold it: the `<t>` elements
 * (directly, or inside `<r>` runs), joined; phonetic runs (`<rPh>`) are
 * skipped wherever they are, with the `<t>` elements inside them (also a
 * run inside a `<t>`: the `<t>`'s text after it is kept). Feed it the
 * events inside the container.
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
    // a <t> inside a phonetic run wasn't counted by open
    else if (name === "t" && this.inPhonetic === 0 && this.inT > 0) this.inT--
  }

  text(value: string): LimitExceeded | undefined {
    if (this.inT === 0 || this.inPhonetic > 0) return undefined
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

/**
 * Reads the shared-string table. An `<si>` inside an `<si>` is
 * MalformedFile: it would replace the open item's text.
 */
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
        if (inItem) return malformed("<si> inside an <si>", entry.name)
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

/** Reads the number formats and the cell styles that use them (`<c s="...">` is an index into `cellXfs`). */
export const readStyles = (
  archive: Archive,
  entry: ZipEntry,
  limits: Limits
): Effect.Effect<CellStyles, SpreadsheetError> => {
  // each custom code is classified once, here in the parse (which yields
  // between pieces), not once per cell style after it: one code is at most
  // maxStringLength characters
  const custom = new Map<number, NumberFormat>()
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
        if (id !== undefined && code !== undefined && /^\d{1,9}$/.test(id)) {
          // a later <numFmt> with the same id wins
          addFormat(custom, Number(id), code)
        }
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
  return Effect.map(parsePart(archive, entry, handler, limits), () => cellStyles(xfIds, xfIds.length, custom))
}

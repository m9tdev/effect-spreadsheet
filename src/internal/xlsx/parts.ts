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
import type { DefinedNameEntry } from "../workbook.ts"
import * as Attr from "../xml/attributes.ts"
import { rulesHandler, type XmlRule } from "../xml/rules.ts"
import type { Archive, LocatedEntry } from "../zip/archive.ts"

export const malformed = malformedIn("xlsx")

/** The largest number with at most nine digits: the bound on the integer attributes read here. */
export const NINE_DIGITS = 999_999_999

export type SheetKind = "worksheet" | "chartsheet" | "dialogsheet" | "macrosheet" | "other"

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
 * open name. The elements are read wherever they sit (a `<sheet>` anywhere
 * inside a `<sheets>`), as wrappers such as `mc:AlternateContent` may hold
 * them.
 */
export const readWorkbookPart = (
  archive: Archive,
  entry: LocatedEntry,
  limits: Limits
): Effect.Effect<WorkbookPart, SpreadsheetError> => {
  // names as the file gives them (decoded; `undefined` for a sheet without one), made unique at the end
  const sheets: Array<Omit<SheetEntry, "name"> & { readonly name: string | undefined }> = []
  const definedNames: Array<DefinedNameEntry> = []
  let date1904 = false
  // the open defined name: its formula is its text, at its end tag
  let name: { name: string; scope: number | undefined; hidden: boolean } | undefined
  const handler = rulesHandler({
    malformed: (reason) => malformed(reason, entry.name),
    forbid: {
      "//sheets//sheets": "<sheets> inside <sheets>",
      "//definedName//definedName": "<definedName> inside a <definedName>"
    },
    on: {
      "//workbookPr": {
        open: (attributes) => {
          date1904 = Attr.bool(attributes, "date1904")
        }
      },
      "//sheets//sheet": {
        open: (attributes) => {
          if (sheets.length >= limits.maxSheets) {
            return new LimitExceeded({ limit: "maxSheets", max: limits.maxSheets, actual: sheets.length + 1 })
          }
          sheets.push({
            name: Attr.xstring(attributes, "name"),
            visibility: sheetVisibility(Attr.string(attributes, "state")),
            relationshipId: Attr.rel(attributes, "id")
          })
          return undefined
        }
      },
      "//definedName": {
        open: (attributes, element) => {
          if (definedNames.length >= limits.maxDefinedNames) {
            return new LimitExceeded({ limit: "maxDefinedNames", max: limits.maxDefinedNames, actual: definedNames.length + 1 })
          }
          // a bad scope is the workbook's, like a missing one
          const local = Attr.uint(attributes, "localSheetId", NINE_DIGITS)
          const scope = local === "bad" ? undefined : local
          name = { name: Attr.xstring(attributes, "name") ?? "", scope, hidden: Attr.bool(attributes, "hidden") }
          // its text is the formula, nested elements' text included
          element.collect()
          return undefined
        },
        close: (text) => {
          // a guard: every <definedName> that gets here was opened by the rule above
          if (name === undefined) return
          definedNames.push({ name: name.name, formula: Attr.unescapeXstring(text ?? ""), hidden: name.hidden, scope: name.scope })
          name = undefined
        }
      }
    }
  })
  return Effect.map(parsePart(archive, entry, handler, limits), () => {
    const names = uniqueSheetNames(sheets.map((s) => s.name))
    return { sheets: sheets.map((s, k) => ({ ...s, name: names[k]! })), definedNames, date1904 }
  })
}

/**
 * The text of a rich string, as `<si>` and `<is>` hold it: the text of
 * its `<t>` elements (directly, or inside `<r>` runs), joined. The text of
 * phonetic runs (`<rPh>`) is left out wherever they are, with the `<t>`
 * elements inside them (also a run inside a `<t>`: the `<t>`'s text after
 * it is kept); their elements still get the rules, so a forbidden one
 * inside them fails. The tokenizer collects each `<t>`'s text, counting
 * the text before it against `maxStringLength`.
 */
export class RichString {
  private parts: Array<string> = []
  private length = 0

  /** The rule for a phonetic run (`<rPh>`) in the string: its text is left out. */
  readonly phonetic: XmlRule = { open: (_, element) => element.omitText() }

  /** The rule for a `<t>` in the string: its text is collected, counting the text before it. */
  readonly text: XmlRule = {
    open: (_, element) => element.collect(this.length),
    // undefined for a <t> inside a <t> (its text is the outer one's) or inside a phonetic run
    close: (text) => {
      if (text === undefined) return
      this.parts.push(text)
      this.length += text.length
    }
  }

  /** The rules for the string's descendants, under the container's path (`//si`). */
  rules(container: string): Record<string, XmlRule> {
    return { [`${container}//rPh`]: this.phonetic, [`${container}//t`]: this.text }
  }

  /** The string's text (unescaped), at the container's end tag; resets for the next one. */
  take(): string {
    const value = Attr.unescapeXstring(this.parts.length === 1 ? this.parts[0]! : this.parts.join(""))
    this.parts = []
    this.length = 0
    return value
  }
}

/**
 * Reads the shared-string table: each `<si>`, wherever it sits. An `<si>`
 * inside an `<si>` is MalformedFile: it would replace the open item's text.
 */
export const readSharedStrings = (
  archive: Archive,
  entry: LocatedEntry,
  limits: Limits
): Effect.Effect<ReadonlyArray<string>, SpreadsheetError> => {
  const strings: Array<string> = []
  const item = new RichString()
  const handler = rulesHandler({
    malformed: (reason) => malformed(reason, entry.name),
    forbid: { "//si//si": "<si> inside an <si>" },
    on: {
      "//si": {
        open: () => {
          if (strings.length >= limits.maxSharedStrings) {
            return new LimitExceeded({ limit: "maxSharedStrings", max: limits.maxSharedStrings, actual: strings.length + 1 })
          }
          return undefined
        },
        close: () => {
          strings.push(item.take())
        }
      },
      ...item.rules("//si")
    }
  })
  return Effect.as(parsePart(archive, entry, handler, limits), strings)
}

/**
 * Reads the number formats and the cell styles that use them (`<c s="...">`
 * is an index into `cellXfs`): each `<numFmt>` anywhere inside a
 * `<numFmts>`, each `<xf>` anywhere inside a `<cellXfs>`.
 */
export const readStyles = (
  archive: Archive,
  entry: LocatedEntry,
  limits: Limits
): Effect.Effect<CellStyles, SpreadsheetError> => {
  // each custom code is classified once, here in the parse (which yields
  // between pieces), not once per cell style after it: one code is at most
  // maxStringLength characters
  const custom = new Map<number, NumberFormat>()
  const xfIds: Array<number> = []
  const handler = rulesHandler({
    malformed: (reason) => malformed(reason, entry.name),
    on: {
      // custom codes only from <numFmts>: <dxfs><dxf><numFmt> (conditional
      // formats) may reuse an id with another code
      "//numFmts//numFmt": {
        open: (attributes) => {
          // one without a valid id or a code is skipped
          const id = Attr.uint(attributes, "numFmtId", NINE_DIGITS)
          const code = Attr.string(attributes, "formatCode")
          // a later <numFmt> with the same id wins
          if (typeof id === "number" && code !== undefined) addFormat(custom, id, code)
        }
      },
      "//cellXfs//xf": {
        open: (attributes) => {
          // a missing or bad id is General (0)
          const id = Attr.uint(attributes, "numFmtId", NINE_DIGITS)
          xfIds.push(typeof id === "number" ? id : 0)
        }
      }
    }
  })
  return Effect.map(parsePart(archive, entry, handler, limits), () => cellStyles(xfIds, xfIds.length, custom))
}

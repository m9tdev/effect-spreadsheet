/**
 * The one rules table for a worksheet part: where its elements sit, which
 * nestings are MalformedFile, and which elements are read only through
 * their own paths. The rows pass and the layout pass (`sheetDetails`) both
 * read the part through it, each giving its own rules for the elements it
 * reads, so they accept and reject the same documents.
 *
 * - `<sheetData>` anywhere. One inside another is MalformedFile: it would
 *   end the sheet data at its end tag, hiding the rows after it (or reading
 *   them as layout). The sheet data is opaque: the layout elements inside
 *   it aren't layout.
 * - `<row>` anywhere inside the sheet data. One inside another is
 *   MalformedFile: it would replace the open row, dropping its cells.
 * - `<c>` anywhere inside a row. A cell is opaque: only its direct `<v>`,
 *   `<f>` and `<is>` children are read, and nothing else in it is a row, a
 *   cell or a nesting error (an `<extLst>` in it is skipped whole). No path
 *   goes on below its `<v>` and `<f>`, so their elements are skipped too,
 *   and only their direct text counts. The `<is>`'s text is rich text: its
 *   `<t>` elements, leaving out phonetic runs (`<rPh>`).
 * - The layout: `<dimension>`, `<col>` and `<mergeCell>` anywhere outside
 *   the sheet data, as wrappers such as `mc:AlternateContent` may hold them.
 *
 * @internal
 */
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { XmlRule, XmlRules } from "../xml/rules.ts"
import { malformed } from "./parts.ts"

const SHEET_DATA = "//sheetData"
const ROW = `${SHEET_DATA}//row`
const CELL = `${ROW}//c`
const VALUE = `${CELL}/v`
const FORMULA = `${CELL}/f`
const INLINE = `${CELL}/is`

/** What one pass over a worksheet part does with the elements it reads; the elements it leaves are matched all the same. */
export interface SheetPass {
  readonly row?: XmlRule
  /** A cell; its `open` may skip it ("skip"), and then its children aren't read. */
  readonly cell?: XmlRule
  /** A cell's direct `<v>`, `<f>` and `<is>` children. */
  readonly value?: XmlRule
  readonly formula?: XmlRule
  readonly inline?: XmlRule
  /** The `<rPh>` and `<t>` elements of an `<is>`. */
  readonly inlinePhonetic?: XmlRule
  readonly inlineText?: XmlRule
  readonly dimension?: XmlRule
  readonly col?: XmlRule
  readonly mergeCell?: XmlRule
}

/** The rules a pass over worksheet `part` reads it with: the shared table, with the pass's rules. */
export const sheetRules = (part: string, pass: SheetPass): XmlRules => ({
  malformed: (reason): SpreadsheetError => malformed(reason, part),
  forbid: {
    [`${SHEET_DATA}//sheetData`]: "<sheetData> inside a <sheetData>",
    [`${ROW}//row`]: "<row> inside a <row>"
  },
  opaque: [SHEET_DATA, CELL],
  on: {
    [ROW]: pass.row ?? {},
    [CELL]: pass.cell ?? {},
    [VALUE]: pass.value ?? {},
    [FORMULA]: pass.formula ?? {},
    [INLINE]: pass.inline ?? {},
    [`${INLINE}//rPh`]: pass.inlinePhonetic ?? {},
    [`${INLINE}//t`]: pass.inlineText ?? {},
    "//dimension": pass.dimension ?? {},
    "//col": pass.col ?? {},
    "//mergeCell": pass.mergeCell ?? {}
  }
})

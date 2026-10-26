/** Builds small XLSX packages for tests (made-up content). */
import { buildZip } from "./zip.ts"

const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const enc = (s: string) => new TextEncoder().encode(s)

export interface TestSheet {
  readonly name: string
  /** The contents of <sheetData> (or a whole part with `raw`). */
  readonly data?: string
  /** A whole worksheet part, instead of `data`. */
  readonly raw?: string
  readonly state?: "hidden" | "veryHidden"
  /** Relationship type, default worksheet. */
  readonly kind?: string
}

export interface TestWorkbook {
  readonly sheets: ReadonlyArray<TestSheet>
  readonly sharedStrings?: ReadonlyArray<string>
  /** Raw sst XML, instead of `sharedStrings`. */
  readonly sharedStringsXml?: string
  /** numFmtId per cellXfs entry, and custom codes. */
  readonly xfs?: ReadonlyArray<number>
  readonly numFmts?: Readonly<Record<number, string>>
  /** Extra styles XML after <cellXfs> (e.g. <dxfs>). */
  readonly stylesExtra?: string
  readonly date1904?: boolean
  readonly definedNames?: string
  /** Extra zip entries. */
  readonly extra?: ReadonlyArray<{ readonly name: string; readonly data: Uint8Array }>
  /** Where the workbook part lives (default xl/workbook.xml). */
  readonly workbookPath?: string
}

const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

export const sheetXml = (data: string) => `<?xml version="1.0" encoding="UTF-8"?><worksheet ${NS}><sheetData>${data}</sheetData></worksheet>`

export const xlsx = (book: TestWorkbook): Promise<Uint8Array> => {
  const wbPath = book.workbookPath ?? "xl/workbook.xml"
  const dir = wbPath.slice(0, wbPath.lastIndexOf("/") + 1)
  const base = wbPath.slice(wbPath.lastIndexOf("/") + 1)
  const sheetsXml = book.sheets
    .map((s, i) => `<sheet name="${escape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"${s.state ? ` state="${s.state}"` : ""}/>`)
    .join("")
  const workbook = `<?xml version="1.0"?><workbook ${NS}>${book.date1904 ? `<workbookPr date1904="1"/>` : ""}<sheets>${sheetsXml}</sheets>${book.definedNames ? `<definedNames>${book.definedNames}</definedNames>` : ""}</workbook>`
  const rels = book.sheets.map((s, i) =>
    `<Relationship Id="rId${i + 1}" Type="${REL}/${s.kind ?? "worksheet"}" Target="worksheets/sheet${i + 1}.xml"/>`
  )
  const entries: Array<{ name: string; data: Uint8Array }> = [
    { name: "_rels/.rels", data: enc(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="${wbPath}"/></Relationships>`) },
    { name: wbPath, data: enc(workbook) }
  ]
  book.sheets.forEach((s, i) => entries.push({ name: `${dir}worksheets/sheet${i + 1}.xml`, data: enc(s.raw ?? sheetXml(s.data ?? "")) }))
  if (book.sharedStrings !== undefined || book.sharedStringsXml !== undefined) {
    rels.push(`<Relationship Id="rIdS" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`)
    const sst = book.sharedStringsXml ??
      `<sst ${NS}>${book.sharedStrings!.map((s) => `<si><t xml:space="preserve">${escape(s)}</t></si>`).join("")}</sst>`
    entries.push({ name: `${dir}sharedStrings.xml`, data: enc(sst) })
  }
  if (book.xfs !== undefined) {
    rels.push(`<Relationship Id="rIdT" Type="${REL}/styles" Target="styles.xml"/>`)
    const fmts = Object.entries(book.numFmts ?? {}).map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${escape(code)}"/>`).join("")
    const styles = `<styleSheet ${NS}><numFmts>${fmts}</numFmts><cellXfs>${book.xfs.map((id) => `<xf numFmtId="${id}"/>`).join("")}</cellXfs>${book.stylesExtra ?? ""}</styleSheet>`
    entries.push({ name: `${dir}styles.xml`, data: enc(styles) })
  }
  entries.push({ name: `${dir}_rels/${base}.rels`, data: enc(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`) })
  for (const e of book.extra ?? []) entries.push(e)
  return buildZip(entries)
}

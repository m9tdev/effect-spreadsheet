/** Builds XLSX and XLSB packages with hand-written relationships (made-up content), for package-layer tests. */
import * as B from "./xlsb.ts"
import { buildZip } from "./zip.ts"

const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`
export const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const enc = (s: string) => new TextEncoder().encode(s)
const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

export interface TestRel {
  readonly id: string
  /** The last segment of the type URI (`worksheet`, `styles`, ...). */
  readonly kind: string
  readonly target: string
  readonly external?: boolean
}

export const relsXml = (rels: ReadonlyArray<TestRel>): Uint8Array =>
  enc(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${
      rels.map((r) =>
        `<Relationship Id="${escape(r.id)}" Type="${REL}/${escape(r.kind)}" Target="${escape(r.target)}"${
          r.external ? ` TargetMode="External"` : ""
        }/>`
      ).join("")
    }</Relationships>`
  )

export interface TestPackage {
  /** Sheet names and relationship ids (none: the sheet has no id), in workbook order. */
  readonly sheets: ReadonlyArray<{ readonly name: string; readonly relationshipId?: string | undefined }>
  /** Where the workbook part is stored. */
  readonly workbookPath: string
  /** The package relationships; `undefined` leaves `_rels/.rels` out. */
  readonly packageRels?: ReadonlyArray<TestRel> | undefined
  /** The workbook part's relationships; `undefined` leaves its rels part out. */
  readonly workbookRels?: ReadonlyArray<TestRel> | undefined
  /** Other parts (worksheets, link parts, ...) by entry name. */
  readonly parts?: ReadonlyArray<{ readonly name: string; readonly data: Uint8Array }>
  /** XLSB only: records after the sheet list (externals, then names). */
  readonly globals?: ReadonlyArray<Uint8Array>
}

/** The relationships part of `part`, as the package names it. */
const relsName = (part: string) => {
  const slash = part.lastIndexOf("/")
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`
}

/** An XLSX package: an XML workbook part, the given relationships and parts. */
export const xlsxPackage = (p: TestPackage): Promise<Uint8Array> => {
  const sheets = p.sheets.map((s, i) => `<sheet name="${escape(s.name)}" sheetId="${i + 1}"${s.relationshipId === undefined ? "" : ` r:id="${escape(s.relationshipId)}"`}/>`).join("")
  const entries: Array<{ name: string; data: Uint8Array }> = []
  if (p.packageRels !== undefined) entries.push({ name: "_rels/.rels", data: relsXml(p.packageRels) })
  entries.push({ name: p.workbookPath, data: enc(`<?xml version="1.0"?><workbook ${NS}><sheets>${sheets}</sheets></workbook>`) })
  if (p.workbookRels !== undefined) entries.push({ name: relsName(p.workbookPath), data: relsXml(p.workbookRels) })
  for (const part of p.parts ?? []) entries.push({ ...part })
  return buildZip(entries)
}

/** An XLSB package: a binary workbook part, the given relationships and parts. */
export const xlsbPackage = (p: TestPackage): Promise<Uint8Array> => {
  const workbook = B.concat([
    B.beginBook(),
    B.wbProp(false),
    B.record(0x8F),
    ...p.sheets.map((s, i) => B.bundleSh(s.name, s.relationshipId, 0, i + 1)),
    B.record(0x90),
    ...(p.globals ?? []),
    B.endBook()
  ])
  const entries: Array<{ name: string; data: Uint8Array }> = []
  if (p.packageRels !== undefined) entries.push({ name: "_rels/.rels", data: relsXml(p.packageRels) })
  entries.push({ name: p.workbookPath, data: workbook })
  if (p.workbookRels !== undefined) entries.push({ name: relsName(p.workbookPath), data: relsXml(p.workbookRels) })
  for (const part of p.parts ?? []) entries.push({ ...part })
  return buildZip(entries)
}

/** A one-cell worksheet part holding `value`, in either format. */
export const worksheet = (format: "xlsx" | "xlsb", value: number): Uint8Array =>
  format === "xlsx"
    ? enc(`<worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>${value}</v></c></row></sheetData></worksheet>`)
    : B.sheetBin([B.rowHdr(0), B.real(0, value)])

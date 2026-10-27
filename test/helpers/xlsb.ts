/** BIFF12 record writers for tests (synthetic content only). */

/** A BIFF12 variable-length integer: 7 bits per byte, low bits first. */
export const varint = (n: number): Array<number> => {
  const out: Array<number> = []
  let rest = n
  do {
    const low = rest % 128
    rest = Math.floor(rest / 128)
    out.push(rest > 0 ? low | 0x80 : low)
  } while (rest > 0)
  return out
}

export const u8 = (n: number) => [n & 0xFF]
export const u16 = (n: number) => [n & 0xFF, (n >>> 8) & 0xFF]
export const u32 = (n: number) => [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF]
export const f64 = (n: number) => {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return [...b]
}

/** An XLWideString: u32 count, UTF-16LE code units. */
export const wide = (s: string): Array<number> => {
  const out = u32(s.length)
  // progress: k up to s.length
  for (let k = 0; k < s.length; k++) out.push(...u16(s.charCodeAt(k)))
  return out
}

/** An XLNullableWideString; `undefined` is the null string. */
export const nullableWide = (s: string | undefined): Array<number> => (s === undefined ? u32(0xFFFFFFFF) : wide(s))

/** One record: type and size varints, then the data. */
export const record = (type: number, data: ReadonlyArray<number> | Uint8Array = []): Uint8Array =>
  new Uint8Array([...varint(type), ...varint(data.length), ...data])

export const concat = (parts: ReadonlyArray<Uint8Array>): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

// ---- records (MS-XLSB record numbers) ----

/** A Cell structure: column u32, style u24 and a flags byte. */
const cellHead = (column: number, style = 0) => [...u32(column), ...u32(style & 0xFFFFFF)]

export const beginBook = () => record(0x83)
export const endBook = () => record(0x84)
export const bundleSh = (name: string, relId: string | undefined, state = 0, tabId = 1) =>
  record(0x9C, [...u32(state), ...u32(tabId), ...nullableWide(relId), ...wide(name)])
export const wbProp = (date1904: boolean) => record(0x99, [...u32(date1904 ? 1 : 0), ...u32(0), ...wide("")])
export const supSelf = () => record(0x165)
export const supAddin = () => record(0x29B)
export const supBookSrc = (relId: string) => record(0x163, wide(relId))
export const externSheet = (xtis: ReadonlyArray<readonly [book: number, first: number, last: number]>) =>
  record(0x16A, [...u32(xtis.length), ...xtis.flatMap(([b, f, l]) => [...u32(b), ...u32(f), ...u32(l)])])
export const supTabs = (names: ReadonlyArray<string>) => record(0x167, [...u32(names.length), ...names.flatMap(wide)])
/** A defined name; `rgce` is its formula's tokens. */
export const name = (n: string, rgce: ReadonlyArray<number>, options: { readonly itab?: number; readonly hidden?: boolean; readonly builtin?: boolean } = {}) =>
  record(0x27, [
    ...u32((options.hidden ? 1 : 0) | (options.builtin ? 0x20 : 0)),
    0,
    ...u32(options.itab ?? 0xFFFFFFFF),
    ...wide(n),
    ...u32(rgce.length),
    ...rgce,
    ...u32(0),
    ...nullableWide(undefined)
  ])
/** PtgRef3d: ixti, row u32, column u16 (bit 14 relative column, bit 15 relative row). */
export const ref3d = (ixti: number, r: number, c: number) => [0x3A, ...u16(ixti), ...u32(r), ...u16(c)]
/** PtgArea3d: ixti, rows u32, columns u16. */
export const area3d = (ixti: number, r1: number, r2: number, c1: number, c2: number) => [0x3B, ...u16(ixti), ...u32(r1), ...u32(r2), ...u16(c1), ...u16(c2)]
export const refErr3d = (ixti: number) => [0x3C, ...u16(ixti), ...new Array(6).fill(0)]
export const areaErr3d = (ixti: number) => [0x3D, ...u16(ixti), ...new Array(12).fill(0)]

export const sstItem = (s: string) => record(0x13, [0, ...wide(s)])
export const beginSst = (count: number) => record(0x9F, [...u32(count), ...u32(count)])
export const fmt = (id: number, code: string) => record(0x2C, [...u16(id), ...wide(code)])
export const xf = (numFmtId: number) => record(0x2F, [...u16(0xFFFF), ...u16(numFmtId), ...new Array(12).fill(0)])

export const beginSheet = () => record(0x81)
export const endSheet = () => record(0x82)
export const beginSheetData = () => record(0x91)
export const endSheetData = () => record(0x92)
/** BrtRowHdr: rw, ixfe, miyRw, flags (fDyZero 0x1000 for hidden), then no column spans. */
export const rowHdr = (r: number, hidden = false) => record(0x00, [...u32(r), ...u32(0), ...u16(300), ...u16(hidden ? 0x1000 : 0), 0, ...u32(0)])
export const blank = (c: number, style = 0) => record(0x01, cellHead(c, style))
export const rk = (c: number, value: number, style = 0) => record(0x02, [...cellHead(c, style), ...u32(value)])
export const cellError = (c: number, code: number) => record(0x03, [...cellHead(c), code])
export const cellBool = (c: number, value: boolean) => record(0x04, [...cellHead(c), value ? 1 : 0])
export const real = (c: number, value: number, style = 0) => record(0x05, [...cellHead(c, style), ...f64(value)])
export const cellSt = (c: number, s: string) => record(0x06, [...cellHead(c), ...wide(s)])
export const isst = (c: number, index: number) => record(0x07, [...cellHead(c), ...u32(index)])
/** A formula cell with its cached result and a one-token formula (PtgInt 1). */
const formulaTail = [...u16(0), ...u32(3), 0x1E, 1, 0, ...u32(0)]
export const fmlaString = (c: number, s: string) => record(0x08, [...cellHead(c), ...wide(s), ...formulaTail])
export const fmlaNum = (c: number, value: number, style = 0) => record(0x09, [...cellHead(c, style), ...f64(value), ...formulaTail])
export const fmlaBool = (c: number, value: boolean) => record(0x0A, [...cellHead(c), value ? 1 : 0, ...formulaTail])
export const fmlaError = (c: number, code: number) => record(0x0B, [...cellHead(c), code, ...formulaTail])
export const rstring = (c: number, s: string) => record(0x3E, [...cellHead(c), 1, ...wide(s), ...u32(1), 0, 0, 0, 0])
/** The short cells (no column: the previous cell's + 1). */
export const shortBlank = (style = 0) => record(0x0C, u32(style))
export const shortRk = (value: number) => record(0x0D, [...u32(0), ...u32(value)])
export const shortReal = (value: number) => record(0x10, [...u32(0), ...f64(value)])
export const shortSt = (s: string) => record(0x11, [...u32(0), ...wide(s)])
export const shortIsst = (index: number) => record(0x12, [...u32(0), ...u32(index)])
export const shortBool = (value: boolean) => record(0x0F, [...u32(0), value ? 1 : 0])
export const shortError = (code: number) => record(0x0E, [...u32(0), code])
const rfx = (r1: number, r2: number, c1: number, c2: number) => [...u32(r1), ...u32(r2), ...u32(c1), ...u32(c2)]
export const wsDim = (r1: number, r2: number, c1: number, c2: number) => record(0x94, rfx(r1, r2, c1, c2))
export const colInfo = (c1: number, c2: number, hidden: boolean) =>
  record(0x3C, [...u32(c1), ...u32(c2), ...u32(2560), ...u32(0), ...u16(hidden ? 1 : 0)])
export const beginColInfos = () => record(0x186)
export const endColInfos = () => record(0x187)
export const mergeCell = (r1: number, r2: number, c1: number, c2: number) => record(0xB0, rfx(r1, r2, c1, c2))
export const beginMergeCells = (count: number) => record(0xB1, u32(count))
export const endMergeCells = () => record(0xB2)

/** RK of an integer (fInt), and of a double's high 30 bits. */
export const rkInt = (n: number) => ((n << 2) | 2) >>> 0
export const rkDouble = (n: number) => new DataView(new Uint8Array(f64(n)).buffer).getUint32(4, true) & 0xFFFFFFFC

/** A worksheet part around `body` (row and cell records), with optional records before and after the sheet data. */
export const sheetBin = (body: ReadonlyArray<Uint8Array>, options: { readonly before?: ReadonlyArray<Uint8Array>; readonly after?: ReadonlyArray<Uint8Array> } = {}) =>
  concat([beginSheet(), ...(options.before ?? []), beginSheetData(), ...body, endSheetData(), ...(options.after ?? []), endSheet()])

// ---- packages ----

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const enc = (s: string) => new TextEncoder().encode(s)

export interface TestSheet {
  readonly name: string
  /** The worksheet part (default: an empty sheet). */
  readonly part?: Uint8Array
  readonly state?: number
  /** Relationship type, default worksheet. */
  readonly kind?: string
}

export interface TestBook {
  readonly sheets: ReadonlyArray<TestSheet>
  readonly strings?: ReadonlyArray<string>
  /** numFmtId per cell XF, and custom codes. */
  readonly xfs?: ReadonlyArray<number>
  readonly formats?: Readonly<Record<number, string>>
  readonly date1904?: boolean
  /** Records after the sheet list (externals, then names). */
  readonly globals?: ReadonlyArray<Uint8Array>
  /** A whole workbook part instead of the generated one. */
  readonly workbookPart?: Uint8Array
  /** External link parts, related from the workbook as rIdX1, rIdX2, ... */
  readonly externalLinks?: ReadonlyArray<Uint8Array>
  readonly extra?: ReadonlyArray<{ readonly name: string; readonly data: Uint8Array }>
  readonly method?: "store" | "deflate"
}

/** An XLSB package (made-up content). */
export const xlsb = async (book: TestBook): Promise<Uint8Array> => {
  const { buildZip } = await import("./zip.ts")
  const workbook = book.workbookPart ?? concat([
    beginBook(),
    wbProp(book.date1904 ?? false),
    record(0x8F),
    ...book.sheets.map((s, i) => bundleSh(s.name, `rId${i + 1}`, s.state ?? 0, i + 1)),
    record(0x90),
    ...(book.globals ?? []),
    endBook()
  ])
  const rels = book.sheets.map((s, i) =>
    `<Relationship Id="rId${i + 1}" Type="${REL}/${s.kind ?? "worksheet"}" Target="worksheets/sheet${i + 1}.bin"/>`
  )
  const entries: Array<{ name: string; data: Uint8Array }> = [
    { name: "[Content_Types].xml", data: enc(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`) },
    { name: "_rels/.rels", data: enc(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.bin"/></Relationships>`) },
    { name: "xl/workbook.bin", data: workbook }
  ]
  book.sheets.forEach((s, i) => entries.push({ name: `xl/worksheets/sheet${i + 1}.bin`, data: s.part ?? sheetBin([]) }))
  if (book.strings !== undefined) {
    rels.push(`<Relationship Id="rIdS" Type="${REL}/sharedStrings" Target="sharedStrings.bin"/>`)
    entries.push({ name: "xl/sharedStrings.bin", data: concat([beginSst(book.strings.length), ...book.strings.map(sstItem), record(0xA0)]) })
  }
  if (book.xfs !== undefined) {
    rels.push(`<Relationship Id="rIdY" Type="${REL}/styles" Target="styles.bin"/>`)
    const formats = Object.entries(book.formats ?? {})
    entries.push({
      name: "xl/styles.bin",
      data: concat([
        record(0x116),
        record(0x267, u32(formats.length)),
        ...formats.map(([id, code]) => fmt(Number(id), code)),
        record(0x268),
        // cell style XFs first: they don't count
        record(0x272, u32(1)),
        xf(0x2A),
        record(0x273),
        record(0x269, u32(book.xfs.length)),
        ...book.xfs.map(xf),
        record(0x26A),
        record(0x117)
      ])
    })
  }
  ;(book.externalLinks ?? []).forEach((part, k) => {
    rels.push(`<Relationship Id="rIdX${k + 1}" Type="${REL}/externalLink" Target="externalLinks/externalLink${k + 1}.bin"/>`)
    entries.push({ name: `xl/externalLinks/externalLink${k + 1}.bin`, data: part })
  })
  entries.push({ name: "xl/_rels/workbook.bin.rels", data: enc(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`) })
  entries.push(...(book.extra ?? []))
  return buildZip(entries.map((e) => ({ ...e, method: book.method ?? "deflate" })))
}

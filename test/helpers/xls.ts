/** Builds small BIFF8 / BIFF5 workbooks for tests (made-up content). */
import { writeCfb } from "./cfb.ts"

const le = (n: number, bytes: number): Array<number> => Array.from({ length: bytes }, (_, k) => Math.floor(n / 2 ** (8 * k)) & 0xFF)
export const u8 = (n: number) => [n & 0xFF]
export const u16 = (n: number) => le(n, 2)
export const u32 = (n: number) => le(n >>> 0, 4)
export const f64 = (n: number) => {
  const b = new Uint8Array(8)
  new DataView(b.buffer).setFloat64(0, n, true)
  return Array.from(b)
}

export const record = (type: number, data: ReadonlyArray<number>): Uint8Array => new Uint8Array([...u16(type), ...u16(data.length), ...data])

const wide = (s: string) => [...s].some((c) => c.charCodeAt(0) > 0xFF) || s.length !== [...s].length
const chars = (s: string, w: boolean) => (w ? Array.from(s, (_, k) => s.charCodeAt(k)).flatMap((c) => u16(c)) : Array.from(s, (_, k) => s.charCodeAt(k) & 0xFF))

/** BIFF8 XLUnicodeString: u16 count, flags, characters (16-bit when needed). */
export const xlString = (s: string) => {
  const w = wide(s)
  return [...u16(s.length), w ? 1 : 0, ...chars(s, w)]
}
/** BIFF8 ShortXLUnicodeString: u8 count, flags, characters. */
export const shortString = (s: string) => {
  const w = wide(s)
  return [s.length, w ? 1 : 0, ...chars(s, w)]
}
/** A BIFF5 string: u8 or u16 count, then bytes (already in the code page). */
export const string5 = (bytes: ReadonlyArray<number>, countBytes: 1 | 2) => [...(countBytes === 1 ? u8(bytes.length) : u16(bytes.length)), ...bytes]

export const BOF = 0x0809
export const bof = (kind: number, biff: 5 | 8 = 8) =>
  record(BOF, biff === 8 ? [...u16(0x0600), ...u16(kind), ...u16(0x0DBB), ...u16(1996), ...u32(0), ...u32(6)] : [...u16(0x0500), ...u16(kind), ...u16(0x0DBB), ...u16(1996)])
export const eof = () => record(0x000A, [])

// globals
export const codepage = (cp: number) => record(0x0042, u16(cp))
export const datemode1904 = () => record(0x0022, u16(1))
export const format = (id: number, code: string, biff: 5 | 8 = 8) =>
  record(0x041E, [...u16(id), ...(biff === 8 ? xlString(code) : string5(Array.from(code, (c) => c.charCodeAt(0)), 1))])
export const xf = (numFmtId: number) => record(0x00E0, [...u16(0), ...u16(numFmtId), ...new Array(16).fill(0)])
export const sst = (strings: ReadonlyArray<string>) => record(0x00FC, [...u32(strings.length), ...u32(strings.length), ...strings.flatMap(xlString)])
export const filepass = () => record(0x002F, [0, 0, 1, 0, ...new Array(50).fill(0)])
/** EXTERNSHEET with entries pointing at sheet indexes (supbook 0). */
export const externSheet = (sheets: ReadonlyArray<number>) => record(0x0017, [...u16(sheets.length), ...sheets.flatMap((s) => [...u16(0), ...u16(s), ...u16(s)])])
/** NAME (BIFF8): `name` (a builtin code when `builtin`), scope sheet index (1-based itab; 0 global), formula tokens. */
export const name = (n: string | number, rgce: ReadonlyArray<number>, options: { itab?: number; hidden?: boolean } = {}) => {
  const builtin = typeof n === "number"
  const text = builtin ? String.fromCharCode(n as number) : (n as string)
  const grbit = (options.hidden ? 0x01 : 0) | (builtin ? 0x20 : 0)
  return record(0x0018, [...u16(grbit), 0, text.length, ...u16(rgce.length), ...u16(0), ...u16(options.itab ?? 0), 0, 0, 0, 0, 0, ...chars(text, false), ...rgce])
}
export const area3d = (ixti: number, r1: number, r2: number, c1: number, c2: number) => [0x3B, ...u16(ixti), ...u16(r1), ...u16(r2), ...u16(c1), ...u16(c2)]
export const ref3d = (ixti: number, r: number, c: number) => [0x3A, ...u16(ixti), ...u16(r), ...u16(c)]

// sheet records
export const dimensions = (r1: number, r2: number, c1: number, c2: number, biff: 5 | 8 = 8) =>
  record(0x0200, biff === 8 ? [...u32(r1), ...u32(r2), ...u16(c1), ...u16(c2), ...u16(0)] : [...u16(r1), ...u16(r2), ...u16(c1), ...u16(c2), ...u16(0)])
export const row = (r: number, hidden = false) => record(0x0208, [...u16(r), ...u16(0), ...u16(1), ...u16(255), ...u16(0), ...u16(0), ...u16(hidden ? 0x20 : 0), ...u16(15)])
const rc = (r: number, c: number, x: number) => [...u16(r), ...u16(c), ...u16(x)]
export const number = (r: number, c: number, x: number, v: number) => record(0x0203, [...rc(r, c, x), ...f64(v)])
export const rk = (r: number, c: number, x: number, value: number) => record(0x027E, [...rc(r, c, x), ...u32(value)])
export const mulrk = (r: number, c: number, cells: ReadonlyArray<readonly [number, number]>) =>
  record(0x00BD, [...u16(r), ...u16(c), ...cells.flatMap(([x, v]) => [...u16(x), ...u32(v)]), ...u16(c + cells.length - 1)])
export const labelSst = (r: number, c: number, x: number, index: number) => record(0x00FD, [...rc(r, c, x), ...u32(index)])
export const label = (r: number, c: number, x: number, s: string) => record(0x0204, [...rc(r, c, x), ...xlString(s)])
export const label5 = (r: number, c: number, x: number, bytes: ReadonlyArray<number>) => record(0x0204, [...rc(r, c, x), ...string5(bytes, 2)])
export const boolErr = (r: number, c: number, x: number, value: number, error: boolean) => record(0x0205, [...rc(r, c, x), value, error ? 1 : 0])
export const blank = (r: number, c: number, x: number) => record(0x0201, rc(r, c, x))
export const mulBlank = (r: number, c: number, xfs: ReadonlyArray<number>) => record(0x00BE, [...u16(r), ...u16(c), ...xfs.flatMap(u16), ...u16(c + xfs.length - 1)])
/** FORMULA with a number result, or a special one: string (then STRING), boolean, error, empty. */
export const formula = (r: number, c: number, x: number, result: number | { kind: 0 | 1 | 2 | 3; value?: number }) => {
  const value = typeof result === "number" ? f64(result) : [result.kind, 0, result.value ?? 0, 0, 0, 0, 0xFF, 0xFF]
  return record(0x0006, [...rc(r, c, x), ...value, ...u16(0), ...u32(0), ...u16(3), 0x1E, 1, 0])
}
export const string = (s: string) => record(0x0207, xlString(s))
export const colInfo = (c1: number, c2: number, hidden: boolean) => record(0x007D, [...u16(c1), ...u16(c2), ...u16(2560), ...u16(0), ...u16(hidden ? 1 : 0), ...u16(0)])
export const mergeCells = (ranges: ReadonlyArray<readonly [number, number, number, number]>) =>
  record(0x00E5, [...u16(ranges.length), ...ranges.flatMap(([r1, r2, c1, c2]) => [...u16(r1), ...u16(r2), ...u16(c1), ...u16(c2)])])

/** RK for a 30-bit integer. */
export const rkInt = (n: number) => ((n << 2) | 2) >>> 0
/** RK for a double whose low 34 bits are zero (e.g. 1.5, -2.25). */
export const rkDouble = (n: number) => new DataView(new Uint8Array(f64(n)).buffer).getUint32(4, true) & 0xFFFFFFFC

export interface TestSheet {
  readonly name: string
  readonly records: ReadonlyArray<Uint8Array>
  /** BOUNDSHEET hsState: 0 visible, 1 hidden, 2 very hidden. */
  readonly state?: number
  /** BOUNDSHEET dt: 0 worksheet, 1 macro, 2 chart, 6 VB module. */
  readonly dt?: number
  /** BOF kind of the substream (default worksheet 0x10). */
  readonly kind?: number
  /** BIFF5 name bytes (in the code page), instead of `name`. */
  readonly name5?: ReadonlyArray<number>
}

export interface TestBook {
  readonly biff?: 5 | 8
  readonly globals?: ReadonlyArray<Uint8Array>
  readonly sheets: ReadonlyArray<TestSheet>
}

const concat = (parts: ReadonlyArray<Uint8Array>) => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

/** The workbook stream: globals (with BOUNDSHEET offsets filled in), then each sheet substream. */
export const workbookStream = (book: TestBook): Uint8Array => {
  const biff = book.biff ?? 8
  const boundsheet = (s: TestSheet, offset: number) =>
    record(0x0085, [...u32(offset), s.state ?? 0, s.dt ?? 0, ...(biff === 8 ? shortString(s.name) : string5(s.name5 ?? Array.from(s.name, (c) => c.charCodeAt(0)), 1))])
  const head = [bof(0x0005, biff), ...(book.globals ?? [])]
  const sheetsBytes = book.sheets.map((s) => concat([bof(s.kind ?? 0x0010, biff), ...s.records, eof()]))
  const globalsLength = concat([...head, ...book.sheets.map((s) => boundsheet(s, 0)), eof()]).length
  let offset = globalsLength
  const bounds = book.sheets.map((s, k) => {
    const r = boundsheet(s, offset)
    offset += sheetsBytes[k]!.length
    return r
  })
  return concat([...head, ...bounds, eof(), ...sheetsBytes])
}

/** A compound file holding the workbook stream (Workbook for BIFF8, Book for BIFF5). */
export const xls = (book: TestBook): Uint8Array =>
  writeCfb([{ name: (book.biff ?? 8) === 8 ? "Workbook" : "Book", data: workbookStream(book) }]).bytes

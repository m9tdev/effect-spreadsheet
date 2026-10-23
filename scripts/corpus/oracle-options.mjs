// Options of the SheetJS oracle (oracle.mjs), in a module without side effects so tests
// and other scripts can import them. Recorded in every expected file.

// 2: rows use sheet_to_json UTC:true and children run with TZ=UTC (schema 1 rows held
// local-time Dates, so they depended on the machine's time zone).
export const SCHEMA = 2
// Options passed to XLSX.read. Recorded in every expected file.
export const READ_OPTIONS = {
  type: "buffer",
  dense: true, // memory scales with present cells, also for huge declared ranges
  cellDates: false, // dates stay serial numbers (v) with their format (z); our reader flags dates itself
  cellNF: true, // keep number format code in z
  cellFormula: true, // keep formulas in f
  cellText: true, // keep SheetJS formatted text in w (for later SSF comparisons)
  cellStyles: true, // needed for XLSX to record hidden rows in !rows (without it the xlsx reader drops the flag)
  sheetStubs: false,
  WTF: false,
}
// sheet_to_json options for `rows` (only computed when the !ref area is small enough).
// UTC: date-formatted cells come back as Dates built in UTC, so `rows` doesn't depend on the
// machine's time zone (their JSON form is an ISO string).
export const JSON_OPTIONS = { header: 1, raw: true, defval: null, blankrows: true, UTC: true }
export const ROWS_MAX_AREA = 2_000_000
// Caps per expected file: above these the file holds status "tooBig" and counts only.
// ORACLE_MAX_CELLS / ORACLE_MAX_BYTES override them (oracle-selftest.mjs uses tiny caps).
const envCap = (k, d) => (process.env[k] && Number(process.env[k]) > 0 ? Number(process.env[k]) : d)
export const MAX_CELLS = envCap("ORACLE_MAX_CELLS", 5_000_000)
export const MAX_BYTES = envCap("ORACLE_MAX_BYTES", 128 * 1024 * 1024)
// A typical SheetJS-based reader: XLSX.read(buffer) then stream.to_json with these options.
export const PROBE_OPTIONS = { raw: true, skipHidden: true, UTC: true, header: "A" }
// Probe variants:
//   default    = XLSX.read(buf) as the old reader did, then stream.to_json(PROBE_OPTIONS)
//   cellStyles = same with cellStyles:true (needed for xlsx to record hidden rows)
//   control    = cellStyles:true but skipHidden:false: if this finishes and the others time out,
//                the timeout is the skipHidden hang, not just a slow range scan.
export const PROBE_VARIANTS = {
  default: { read: { type: "buffer" }, stream: PROBE_OPTIONS },
  cellStyles: { read: { type: "buffer", cellStyles: true }, stream: PROBE_OPTIONS },
  control: { read: { type: "buffer", cellStyles: true }, stream: { ...PROBE_OPTIONS, skipHidden: false } },
}


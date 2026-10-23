/**
 * Resource limits applied to every read. All input is untrusted, so every
 * size, count and depth a reader meets is checked against these before use;
 * going over fails with `LimitExceeded`.
 *
 * Defaults are safe for untrusted uploads in a browser tab or a server
 * worker. Override them by layer (`SpreadsheetLimits.layer({...})`) or per
 * call (`{ limits: {...} }` on any reader's options); per-call values win.
 *
 * @since 0.1.0
 */
import { Context, Effect, Layer } from "effect"

/**
 * @since 0.1.0
 * @category models
 */
export interface Limits {
  /** Size of the input file in bytes (for streams: bytes read so far). */
  readonly maxInputBytes: number
  /** Total bytes inflated from a container (all zip entries together). */
  readonly maxDecompressedBytes: number
  /**
   * Max inflated/compressed ratio of a zip entry. Checked once an entry has
   * inflated past 1 MiB, so tiny highly-compressible entries pass.
   */
  readonly maxCompressionRatio: number
  /** Entries in a zip central directory. */
  readonly maxZipEntries: number
  /** Sheets in a workbook. */
  readonly maxSheets: number
  /** Rows per sheet: row indexes must be `< maxRows`. */
  readonly maxRows: number
  /** Columns per row: column indexes must be `< maxColumns`. */
  readonly maxColumns: number
  /** Cells actually present (non-empty) in one sheet. */
  readonly maxCellsPerSheet: number
  /** Length of one string value (cell text, shared string), in UTF-16 code units. */
  readonly maxStringLength: number
  /** Entries in a shared-string table. */
  readonly maxSharedStrings: number
  /** Element nesting depth of an XML part. */
  readonly maxXmlDepth: number
  /** Attributes on one XML element. */
  readonly maxXmlAttributes: number
  /** Sectors in one CFB (OLE) FAT/miniFAT chain. */
  readonly maxSectorChain: number
}

/**
 * Some limits to override; the others are inherited. `undefined` means
 * "inherit" too.
 *
 * @since 0.1.0
 * @category models
 */
export type LimitOverrides = { readonly [K in keyof Limits]?: number | undefined }

/**
 * The defaults:
 *
 * | limit | default | why |
 * |---|---|---|
 * | maxInputBytes | 256 MiB | far above real uploads, below tab-crashing sizes |
 * | maxDecompressedBytes | 1 GiB | a 200k-row XLSX inflates to a few hundred MiB |
 * | maxCompressionRatio | 100 | the Apache POI zip-bomb default (1 / MIN_INFLATE_RATIO 0.01) |
 * | maxZipEntries | 10 000 | real workbooks have tens to hundreds |
 * | maxSheets | 1 024 | |
 * | maxRows | 1 048 576 | Excel's row limit |
 * | maxColumns | 16 384 | Excel's column limit (XFD) |
 * | maxCellsPerSheet | 20 000 000 | about 1M rows x 20 columns |
 * | maxStringLength | 262 144 | 8x Excel's 32 767-character cell limit, room for CSV blobs |
 * | maxSharedStrings | 5 000 000 | |
 * | maxXmlDepth | 128 | worksheet XML is about 6 levels deep |
 * | maxXmlAttributes | 512 | |
 * | maxSectorChain | 1 048 576 | 512 MiB of 512-byte sectors |
 *
 * @since 0.1.0
 * @category defaults
 */
export const defaults: Limits = {
  maxInputBytes: 256 * 1024 * 1024,
  maxDecompressedBytes: 1024 * 1024 * 1024,
  maxCompressionRatio: 100,
  maxZipEntries: 10_000,
  maxSheets: 1_024,
  maxRows: 1_048_576,
  maxColumns: 16_384,
  maxCellsPerSheet: 20_000_000,
  maxStringLength: 262_144,
  maxSharedStrings: 5_000_000,
  maxXmlDepth: 128,
  maxXmlAttributes: 512,
  maxSectorChain: 1_048_576
}

/**
 * The limits in effect. A `Context.Reference`, so it needs no layer: without
 * one it is {@link defaults}.
 *
 * @since 0.1.0
 * @category references
 */
export const SpreadsheetLimits: Context.Reference<Limits> = Context.Reference<Limits>(
  "effect-spreadsheet/SpreadsheetLimits",
  { defaultValue: () => defaults }
)

/**
 * A layer that overrides some limits; the others keep their defaults.
 *
 * @since 0.1.0
 * @category layers
 */
export const layer = (overrides: LimitOverrides): Layer.Layer<never> =>
  Layer.succeed(SpreadsheetLimits, merge(defaults, overrides))

/**
 * @since 0.1.0
 * @category combinators
 */
export const merge = (base: Limits, overrides: LimitOverrides | undefined): Limits => {
  if (overrides === undefined) return base
  const out: { -readonly [K in keyof Limits]: number } = { ...base }
  for (const key of Object.keys(base) as Array<keyof Limits>) {
    const value = overrides[key]
    if (value !== undefined) out[key] = value
  }
  return out
}

/**
 * Combines two sets of overrides; defined values of `overrides` win.
 *
 * @since 0.1.0
 * @category combinators
 */
export const mergePartial = (
  base: LimitOverrides | undefined,
  overrides: LimitOverrides | undefined
): LimitOverrides | undefined => {
  if (overrides === undefined) return base
  if (base === undefined) return overrides
  const out: { -readonly [K in keyof Limits]?: number | undefined } = { ...base }
  for (const key of Object.keys(overrides) as Array<keyof Limits>) {
    const value = overrides[key]
    if (value !== undefined) out[key] = value
  }
  return out
}

/**
 * The limits for one call: the context's limits with the call's overrides on
 * top.
 *
 * @since 0.1.0
 * @category accessors
 */
export const resolve = (overrides?: LimitOverrides): Effect.Effect<Limits> =>
  Effect.map(SpreadsheetLimits, (limits) => merge(limits, overrides))

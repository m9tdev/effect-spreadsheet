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
  /** Total bytes inflated from a container (all zip entries together): the hard cap. */
  readonly maxDecompressedBytes: number
  /**
   * Archive-wide zip-bomb guard: all entries together may inflate to at
   * most `max(inflateFloorBytes, maxInflateRatio × archive size)` bytes
   * (and never more than `maxDecompressedBytes`). Per archive, not per
   * entry, so many small highly compressible entries can't add up.
   */
  readonly maxInflateRatio: number
  /** The inflate budget every archive gets, however small it is (see `maxInflateRatio`). */
  readonly inflateFloorBytes: number
  /** Entries in a zip central directory. */
  readonly maxZipEntries: number
  /** Sheets in a workbook. */
  readonly maxSheets: number
  /** Rows per sheet: row indexes must be `< maxRows`. */
  readonly maxRows: number
  /** Columns per row: column indexes must be `< maxColumns`. */
  readonly maxColumns: number
  /**
   * Cells emitted for one sheet: every cell in every `Row.cells`, `Empty`
   * cells included (a styled blank in XLSX counts). Positions without a
   * cell (sparse gaps, CSV empty fields) don't count.
   */
  readonly maxCellsPerSheet: number
  /**
   * `defval` fills written by `Spreadsheet.toRecords` in one run (keys
   * added for header columns a row has no value in). Smaller than
   * `maxCellsPerSheet`, since a wide header times many short rows would
   * otherwise cost far more than the cells present.
   */
  readonly maxRecordFields: number
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
  /** Nesting depth of storages in a CFB container (a workbook's are 1-3 deep). */
  readonly maxStorageDepth: number
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
 * | maxInflateRatio | 200 | real workbooks inflate 5-20x; blank-row sheets reach a few hundred |
 * | inflateFloorBytes | 64 MiB | small files may inflate this far whatever their size |
 * | maxZipEntries | 10 000 | real workbooks have tens to hundreds |
 * | maxSheets | 1 024 | |
 * | maxRows | 1 048 576 | Excel's row limit |
 * | maxColumns | 16 384 | Excel's column limit (XFD) |
 * | maxCellsPerSheet | 20 000 000 | about 1M rows x 20 columns |
 * | maxRecordFields | 1 048 576 | defval fills per toRecords run: ~1M keys, well under a second |
 * | maxStringLength | 262 144 | 8x Excel's 32 767-character cell limit, room for CSV blobs |
 * | maxSharedStrings | 5 000 000 | |
 * | maxXmlDepth | 128 | worksheet XML is about 6 levels deep |
 * | maxXmlAttributes | 512 | |
 * | maxSectorChain | 1 048 576 | 512 MiB of 512-byte sectors |
 * | maxStorageDepth | 32 | VBA projects nest storages 2-3 deep; paths cost depth² |
 *
 * @since 0.1.0
 * @category defaults
 */
export const defaults: Limits = {
  maxInputBytes: 256 * 1024 * 1024,
  maxDecompressedBytes: 1024 * 1024 * 1024,
  maxInflateRatio: 200,
  inflateFloorBytes: 64 * 1024 * 1024,
  maxZipEntries: 10_000,
  maxSheets: 1_024,
  maxRows: 1_048_576,
  maxColumns: 16_384,
  maxCellsPerSheet: 20_000_000,
  maxRecordFields: 1_048_576,
  maxStringLength: 262_144,
  maxSharedStrings: 5_000_000,
  maxXmlDepth: 128,
  maxXmlAttributes: 512,
  maxSectorChain: 1_048_576,
  maxStorageDepth: 32
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
 * An invalid value (see {@link validate}) makes building the layer die.
 *
 * @since 0.1.0
 * @category layers
 */
export const layer = (overrides: LimitOverrides): Layer.Layer<never> =>
  Layer.sync(SpreadsheetLimits)(() => merge(defaults, overrides))

/**
 * A valid limit is a non-negative integer or `Infinity` (no limit). `NaN`,
 * negative, fractional or non-number values would silently disable checks
 * like `index >= max`, so they are programmer errors.
 *
 * @since 0.1.0
 * @category guards
 */
export const isValidLimit = (value: unknown): value is number =>
  value === Infinity || (typeof value === "number" && Number.isInteger(value) && value >= 0)

/**
 * Checks every limit; throws a `TypeError` naming the first invalid one.
 *
 * @since 0.1.0
 * @category guards
 */
export const validate = (limits: Limits): Limits => {
  for (const key of Object.keys(defaults) as Array<keyof Limits>) {
    const value: unknown = limits[key]
    if (!isValidLimit(value)) {
      throw new TypeError(
        `SpreadsheetLimits.${key} must be a non-negative integer or Infinity, got ${typeof value === "number" ? value : typeof value}`
      )
    }
  }
  return limits
}

/**
 * `base` with the defined `overrides` on top, validated: throws a
 * `TypeError` if any resulting limit is invalid (see {@link validate}).
 *
 * @since 0.1.0
 * @category combinators
 */
export const merge = (base: Limits, overrides: LimitOverrides | undefined): Limits => {
  const out: { -readonly [K in keyof Limits]: number } = { ...base }
  if (overrides !== undefined) {
    for (const key of Object.keys(defaults) as Array<keyof Limits>) {
      const value = overrides[key]
      if (value !== undefined) out[key] = value
    }
  }
  return validate(out)
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
 * top. Invalid values (in the context or the overrides) make it die.
 *
 * @since 0.1.0
 * @category accessors
 */
export const resolve = (overrides?: LimitOverrides): Effect.Effect<Limits> =>
  Effect.flatMap(SpreadsheetLimits, (limits) => Effect.sync(() => merge(limits, overrides)))

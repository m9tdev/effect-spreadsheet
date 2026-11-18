/**
 * Declared record layouts for the binary formats: a record's data is
 * checked once against the minimum length its layout declares, and its
 * fixed fields are then read without a check (and without `undefined`) per
 * field; and the one result type of a counted string read.
 *
 * `FixedReader.fixed` allocates nothing: a reader is made once per part
 * read and pointed at each record in turn, and its fields are read from
 * the record's bytes directly (a float through an 8-byte scratch buffer),
 * not through a `DataView` made per record.
 *
 * @internal
 */

/**
 * The fixed fields of a record whose length was checked: little-endian
 * reads at absolute offsets below the checked minimum. Offsets past it
 * aren't checked here; a variable-length tail checks its own.
 */
export interface Fixed {
  readonly data: Uint8Array
  u8(at: number): number
  u16(at: number): number
  u32(at: number): number
  i32(at: number): number
  f64(at: number): number
}

const EMPTY = new Uint8Array(0)

/** The scratch bytes a float is assembled in: one shared buffer, read as soon as it is written. */
const SCRATCH = new Uint8Array(8)
const SCRATCH_VIEW = new DataView(SCRATCH.buffer)

/** One reader for a part's records, pointed at each record by `fixed`. */
export class FixedReader implements Fixed {
  data: Uint8Array = EMPTY

  /**
   * `data` with its fixed fields, if it holds at least `min` bytes (the
   * record's declared minimum); `undefined` if it is shorter: the one
   * "truncated" check. Repoints this reader, so a `Fixed` it gave before
   * reads the new record.
   */
  fixed(data: Uint8Array, min: number): Fixed | undefined {
    if (data.length < min) return undefined
    this.data = data
    return this
  }

  u8(at: number): number {
    return this.data[at]!
  }

  u16(at: number): number {
    const d = this.data
    return d[at]! | (d[at + 1]! << 8)
  }

  u32(at: number): number {
    return this.i32(at) >>> 0
  }

  i32(at: number): number {
    const d = this.data
    return d[at]! | (d[at + 1]! << 8) | (d[at + 2]! << 16) | (d[at + 3]! << 24)
  }

  f64(at: number): number {
    const d = this.data
    // progress: k up to 8
    for (let k = 0; k < 8; k++) SCRATCH[k] = d[at + k]!
    return SCRATCH_VIEW.getFloat64(0, true)
  }
}

/**
 * A counted string read at some offset: its text and the offset after it;
 * or it runs past the data (a broken record, whatever the limit); or its
 * length is over the limit, with that length (for `LimitExceeded.actual`).
 * Both failures are found before anything is decoded.
 */
export type StringRead<A = string> =
  | { readonly _tag: "Ok"; readonly value: A; readonly end: number }
  | { readonly _tag: "Truncated" }
  | { readonly _tag: "TooLong"; readonly count: number }

/** A string read that failed. */
export type StringFailure = Exclude<StringRead<unknown>, { readonly _tag: "Ok" }>

export const TRUNCATED: StringFailure = { _tag: "Truncated" }

export const ok = <A>(value: A, end: number): StringRead<A> => ({ _tag: "Ok", value, end })

export const tooLongString = (count: number): StringFailure => ({ _tag: "TooLong", count })

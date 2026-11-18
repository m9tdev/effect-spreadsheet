/**
 * Fixed-layout reads shared by the binary formats (XLS and XLSB records):
 * bounds-checked fields at absolute offsets, RK numbers and error codes.
 *
 * @internal
 */

/** Reads at absolute offsets of one record's data; out of bounds gives `undefined`. */
export class Fields {
  readonly data: Uint8Array
  private readonly view: DataView

  constructor(data: Uint8Array) {
    this.data = data
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  }

  get length(): number {
    return this.data.length
  }

  u8(at: number): number | undefined {
    return at >= 0 && at + 1 <= this.data.length ? this.data[at] : undefined
  }

  u16(at: number): number | undefined {
    return at >= 0 && at + 2 <= this.data.length ? this.view.getUint16(at, true) : undefined
  }

  u32(at: number): number | undefined {
    return at >= 0 && at + 4 <= this.data.length ? this.view.getUint32(at, true) : undefined
  }

  f64(at: number): number | undefined {
    return at >= 0 && at + 8 <= this.data.length ? this.view.getFloat64(at, true) : undefined
  }
}

/** Excel error codes (BOOLERR, FORMULA results). */
export const ERROR_CODES: ReadonlyMap<number, string> = new Map([
  [0x00, "#NULL!"],
  [0x07, "#DIV/0!"],
  [0x0F, "#VALUE!"],
  [0x17, "#REF!"],
  [0x1D, "#NAME?"],
  [0x24, "#NUM!"],
  [0x2A, "#N/A"],
  [0x2B, "#GETTING_DATA"]
])

/** An RK number: a 30-bit integer or the high 30 bits of a double, optionally divided by 100. */
export const decodeRk = (rk: number): number => {
  let value: number
  if ((rk & 0x02) !== 0) value = rk >> 2
  else {
    const buffer = new DataView(new ArrayBuffer(8))
    buffer.setUint32(4, rk & 0xFFFFFFFC, true)
    buffer.setUint32(0, 0, true)
    value = buffer.getFloat64(0, true)
  }
  return (rk & 0x01) !== 0 ? value / 100 : value
}

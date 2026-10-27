/** A synthetic BIFF8 workbook with many rows (made-up content), for the performance checks. */
import * as X from "./xls.ts"

/**
 * `rows` rows (at most 65 536, BIFF8's limit) of 4 cells: a shared string,
 * a number, an RK integer and a boolean; ROW records in blocks of 32 before
 * their cells, every 10th row hidden.
 */
export const bigXls = (rows: number): Uint8Array => {
  const records: Array<Uint8Array> = []
  for (let block = 0; block < rows; block += 32) {
    const end = Math.min(rows, block + 32)
    for (let r = block; r < end; r++) records.push(X.row(r, r % 10 === 9))
    for (let r = block; r < end; r++) {
      records.push(X.labelSst(r, 0, 0, r % 3), X.number(r, 1, 0, r * 1.5), X.rk(r, 2, 0, X.rkInt(r)), X.boolErr(r, 3, 0, r & 1, false))
    }
  }
  return X.xls({ globals: [X.xf(0), X.sst(["alpha", "beta", "gamma"])], sheets: [{ name: "Big", records }] })
}

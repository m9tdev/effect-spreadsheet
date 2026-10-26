/** A synthetic workbook with many rows (made-up content), for the performance checks. */
import { xlsx } from "./xlsx.ts"

/** Bytes of the shortest row `bigXlsx` writes (row 1), a lower bound for every row. */
export const MIN_ROW_BYTES = 120

/**
 * `rows` rows of 4 cells: a shared string, a number, a date and an inline
 * string; every 10th row hidden.
 */
export const bigXlsx = (rows: number): Promise<Uint8Array> => {
  const parts: Array<string> = []
  for (let i = 1; i <= rows; i++) {
    parts.push(
      `<row r="${i}"${i % 10 === 0 ? ` hidden="1"` : ""}><c r="A${i}" t="s"><v>${i % 3}</v></c><c r="B${i}"><v>${i * 1.5}</v></c>` +
        `<c r="C${i}" s="1"><v>${46000 + (i % 1000)}</v></c><c r="D${i}" t="inlineStr"><is><t>row ${i}</t></is></c></row>`
    )
  }
  return xlsx({ sharedStrings: ["alpha", "beta", "gamma"], xfs: [0, 14], sheets: [{ name: "Big", data: parts.join("") }] })
}

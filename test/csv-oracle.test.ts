/**
 * CSV vs the SheetJS oracle (~/sheetjs/xlsx.mjs; skipped when absent):
 * our rows against `XLSX.read(csv, { type: "buffer" })` +
 * `sheet_to_json(ws, { header: 1, raw: true })`.
 *
 * - inferTypes: false is compared with SheetJS `read({ raw: true })` (all text);
 * - inferTypes: true with SheetJS's default (inferring) read.
 *
 * The intended differences are listed (and checked) below; see README.
 */
import { describe, expect, it } from "vitest"
import { csvAoa, enc } from "./helpers/run.ts"
import { aoa, load } from "./oracle/sheetjs.ts"

const xlsx = await load()

const SAME: Record<string, string> = {
  simple: "a,b,c\n1,2,3\n",
  quoted: '"a,b","c""d","e\nf"\nx,"",z\n',
  crlf: "a,b\r\nc,d\r\n",
  cr: "a,b\rc,d\r",
  noTrailingBreak: "a,b\nc,d",
  semicolon: "a;b;c\n1;2;3\n",
  tab: "a\tb\n1\t2\n",
  pipe: "a|b\n1|2\n",
  ragged: "a\n1,2,3\n,,x\n",
  blankLines: "a\n\nb\n",
  trailingBlankLines: "a\nb\n\n\n",
  sepLine: "sep=;\na;b\n",
  sepLineCrlf: "sep=|\r\na|b\r\n",
  quoteInsideUnquoted: 'ab"c,d"\n',
  utf8Bom: "\uFEFFnév,ár\nÉva,5\n",
  numbers: "1,-2.5,.5,1e3,+7,0,0.25,TRUE,FALSE\n",
  wide: Array.from({ length: 300 }, (_, i) => `c${i}`).join(",") + "\n",
  many: Array.from({ length: 2000 }, (_, i) => `${i},row ${i},"q,${i}"`).join("\r\n") + "\r\n"
}

describe.skipIf(xlsx === undefined)("csv vs SheetJS oracle", () => {
  for (const [name, text] of Object.entries(SAME)) {
    it(`same rows: ${name}`, async () => {
      const bytes = enc(text)
      expect(await csvAoa(bytes)).toEqual(aoa(xlsx!, bytes, { raw: true }))
      expect(await csvAoa(bytes, { inferTypes: true })).toEqual(aoa(xlsx!, bytes))
    })
  }

  // Intended differences: [input, ours (inferTypes: true), SheetJS default read]
  const DIFFERENT: Record<string, readonly [string, unknown, unknown]> = {
    // Excel-like: the closing quote ends the quoted part, the rest is appended. SheetJS keeps the raw text.
    textAfterClosingQuote: ['"ab"c,d\n', [["abc", "d"]], [['"ab"c', "d"]]],
    // An unterminated quote runs to EOF in both; SheetJS keeps the opening quote.
    unterminatedQuote: ['a,"b\nc\n', [["a", "b\nc\n"]], [["a", '"b\nc\n']]],
    // UTF-8 without BOM: we detect UTF-8; SheetJS (type "buffer") reads bytes as Latin-1.
    utf8WithoutBom: ["név\n", [["név"]], [["nÃ©v"]]],
    // Leading zeros stay text (zip codes, IDs); SheetJS gives 7.
    leadingZeros: ["007\n", [["007"]], [[7]]],
    // More than 15 significant digits stay text (no precision loss).
    longDigits: ["1234567890123456789\n", [["1234567890123456789"]], [[1234567890123456800]]],
    // Booleans in any case; SheetJS only TRUE/FALSE.
    lowercaseBoolean: ["true\n", [[true]], [["true"]]],
    // No fuzzy numbers: %, currency, padding stay text.
    fuzzyNumbers: ['50%,$5," 1"\n', [["50%", "$5", " 1"]], [[0.5, 5, 1]]],
    // No date inference (dd/mm vs mm/dd is ambiguous); SheetJS gives a serial.
    dates: ["2024-01-02\n", [["2024-01-02"]], [[45293]]],
    // Error literals are Error cells with their code; SheetJS raw sheet_to_json gives null/undefined.
    errorLiteral: ["#N/A,x\n", [["#N/A", "x"]], [[undefined, "x"]]],
    // A file starting with a space: SheetJS takes a different (PRN) path and returns byte values.
    leadingSpace: [" a , b \n", [[" a ", " b "]], [[32, 97, 32, 44, 32, 98, 32, 10]]]
  }

  for (const [name, [text, ours, theirs]] of Object.entries(DIFFERENT)) {
    it(`intended difference: ${name}`, async () => {
      const bytes = enc(text)
      expect(await csvAoa(bytes, { inferTypes: true })).toEqual(ours)
      expect(aoa(xlsx!, bytes)).toEqual(theirs)
    })
  }
})

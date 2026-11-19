/**
 * Times the CSV hot loops on their own, without decoding or streams: the
 * parser over a whole text fed in 64 Ki-character pieces, with the
 * delimiter given, and the delimiter guess over 2000 16 Ki-character samples.
 * Several input shapes (plain, quoted, wide, short lines, tab-delimited).
 * Also many tiny inputs, where the cost of setting up a read shows: the
 * parser with the delimiter given or guessed, the guess alone, and the
 * public `Csv.rows` end to end.
 * Prints one JSON line: the median milliseconds of each case.
 * Not a test: the numbers depend on the machine and its load; compare two
 * builds by alternating their runs. Run: node scripts/bench/csv-lex.ts [rounds]
 */
import { Effect, Stream } from "effect"
import * as Result from "effect/Result"
import * as Csv from "../../src/csv/index.ts"
import { CsvParser } from "../../src/internal/csv/parser.ts"
import { guessDelimiter, SNIFF_CHARS } from "../../src/internal/csv/sniff.ts"
import * as SpreadsheetLimits from "../../src/SpreadsheetLimits.ts"

const rounds = Number(process.argv[2] ?? 15)

const lines = (count: number, line: (i: number) => string, eol = "\r\n"): string => {
  const out: Array<string> = []
  // progress: i increases up to count
  for (let i = 0; i < count; i++) out.push(line(i))
  return out.join(eol) + eol
}

const SHAPES: Record<string, { readonly text: string; readonly delimiter: string }> = {
  plain: { text: lines(200_000, (i) => `${i},name ${i % 997},city ${i % 31},${(i * 1.25).toFixed(2)},2024-01-0${i % 9 + 1}`), delimiter: "," },
  quoted: {
    text: lines(200_000, (i) => `${i},"Name ${i % 997}, Jr.","said ""hi"" ${i % 7}","line\nbreak",Zürich`),
    delimiter: ","
  },
  wide: { text: lines(10_000, (i) => Array.from({ length: 100 }, (_, c) => `${i}.${c}`).join(";"), "\n"), delimiter: ";" },
  shortLines: { text: lines(1_000_000, (i) => `${i % 10},x`, "\n"), delimiter: "," },
  tabs: { text: lines(200_000, (i) => `${i}\tsome longer text in a field ${i}\t${i % 3}\t\t`), delimiter: "\t" }
}

const PIECE = 64 * 1024

const parse = (text: string, delimiter: string): number => {
  const parser = new CsvParser({
    delimiter: delimiter.charCodeAt(0),
    inferTypes: false,
    strict: false,
    limits: SpreadsheetLimits.defaults
  })
  let rows = 0
  // progress: i increases by PIECE up to text.length
  for (let i = 0; i < text.length; i += PIECE) {
    const r = parser.feed(text.slice(i, i + PIECE), i + PIECE >= text.length)
    if (Result.isFailure(r)) throw new Error(r.failure._tag)
    rows += r.success.length
  }
  return rows
}

const SAMPLES = 2000

/**
 * Guesses the delimiter of SAMPLES 16 Ki-character samples of the text, in
 * turn complete or not (cycling through the text: a guess stops after 64
 * lines, so short lines make quick guesses).
 */
const sniff = (text: string): number => {
  const samples: Array<string> = []
  // progress: i increases by SNIFF_CHARS up to text.length
  for (let i = 0; i < text.length; i += SNIFF_CHARS) samples.push(text.slice(i, i + SNIFF_CHARS))
  let sum = 0
  // progress: k increases up to SAMPLES
  for (let k = 0; k < SAMPLES; k++) sum += guessDelimiter(samples[k % samples.length]!, k % 2 === 0)
  return sum
}

const median = (xs: Array<number>): number => {
  const s = [...xs].sort((a, b) => a - b)
  return s[s.length >> 1]!
}

const time = (f: () => number): Array<number> => {
  const ms: Array<number> = []
  // progress: r increases up to rounds
  for (let r = 0; r < rounds; r++) {
    const start = performance.now()
    f()
    ms.push(performance.now() - start)
  }
  return ms
}

const TINY = "a,b,c\n1,2,3\n"
const TINY_BYTES = new TextEncoder().encode(TINY)

/** Parses TINY `count` times, each with a new parser. */
const tinyParse = (delimiter: number | undefined, count: number) => (): number => {
  let rows = 0
  // progress: k increases up to count
  for (let k = 0; k < count; k++) {
    const parser = new CsvParser({ delimiter, inferTypes: false, strict: false, limits: SpreadsheetLimits.defaults })
    const r = parser.feed(TINY, true)
    if (Result.isFailure(r)) throw new Error(r.failure._tag)
    rows += r.success.length
  }
  return rows
}

const tinySniff = (count: number) => (): number => {
  let sum = 0
  // progress: k increases up to count
  for (let k = 0; k < count; k++) sum += guessDelimiter(TINY, true)
  return sum
}

const tinyRows = (count: number) => (): number => {
  let rows = 0
  // progress: k increases up to count
  for (let k = 0; k < count; k++) rows += Effect.runSync(Stream.runCollect(Csv.rows(TINY_BYTES))).length
  return rows
}

const ms = (f: () => number): number => Math.round(median(time(f)) * 10) / 10

const result: Record<string, number> = {}
for (const [name, { delimiter, text }] of Object.entries(SHAPES)) {
  result[`parse.${name}`] = ms(() => parse(text, delimiter))
  result[`sniff.${name}`] = ms(() => sniff(text))
}
result["tiny.parse.known"] = ms(tinyParse(0x2C, 100_000))
result["tiny.parse.sniffed"] = ms(tinyParse(undefined, 100_000))
result["tiny.sniff"] = ms(tinySniff(300_000))
result["tiny.public.rows"] = ms(tinyRows(5_000))
console.log(JSON.stringify(result))

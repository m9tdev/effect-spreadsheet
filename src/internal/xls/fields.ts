/**
 * The string forms BIFF8 and BIFF5 records use, and the text decoding of
 * the workbook's code page.
 *
 * @internal
 */
import { compressed, readRichString, SegmentReader, segmentsOf, utf16 } from "../biff/segments.ts"

/**
 * Text decoding for the workbook's code page (CODEPAGE record). BIFF5
 * strings are bytes in this code page; BIFF8 strings are UTF-16 (or its
 * low bytes) whatever it says. Code pages `TextDecoder` doesn't know (the
 * DOS ones, 437 and 850, are the common case) fall back to windows-1252.
 */
export const codepageLabel = (codepage: number): string => {
  if (codepage >= 1250 && codepage <= 1258) return `windows-${codepage}`
  if (codepage >= 28591 && codepage <= 28606) return `iso-8859-${codepage - 28590}`
  switch (codepage) {
    case 1200: return "utf-16le"
    case 1201: return "utf-16be"
    case 65001: return "utf-8"
    case 874: return "windows-874"
    case 932: return "shift_jis"
    case 936: return "gbk"
    case 949: return "euc-kr"
    case 950: return "big5"
    case 866: return "ibm866"
    case 20866: return "koi8-r"
    case 21866: return "koi8-u"
    case 10000:
    case 32768: return "macintosh"
    default: return "windows-1252"
  }
}

/**
 * A decoder for an encoding label, or windows-1252 when the runtime doesn't
 * have that encoding (`TextDecoder` throws on unknown labels; runtimes
 * without full ICU lack the East Asian ones).
 */
export const textDecoder = (label: string): TextDecoder => {
  try {
    return new TextDecoder(label)
  } catch {
    return new TextDecoder("windows-1252")
  }
}

export interface Text {
  /** BIFF8 (UTF-16 strings) or BIFF5 (code-page strings). */
  readonly biff8: boolean
  readonly decoder: TextDecoder
  readonly maxLength: number
}

/**
 * A BIFF8 XLUnicodeString (u16 count, flags, characters, optional run and
 * phonetic data) or a BIFF5 string (u16 or u8 count, code-page bytes) at
 * `at`. Returns the text and the offset after it; `undefined` if it runs
 * past the data, `null` if it is longer than the limit.
 */
export const readString = (
  data: Uint8Array,
  at: number,
  text: Text,
  countBytes: 1 | 2
): { readonly value: string; readonly end: number } | undefined | null => {
  if (at > data.length) return undefined
  const rest = data.subarray(at)
  if (text.biff8) {
    if (countBytes === 2) {
      const reader = new SegmentReader(segmentsOf([rest]))
      const value = readRichString(reader, text.maxLength)
      if (value === undefined || value === null) return value
      return { value, end: data.length - reader.remaining() }
    }
    // ShortXLUnicodeString: u8 count, flags, characters
    if (rest.length < 2) return undefined
    const count = rest[0]!
    if (count > text.maxLength) return null
    const wide = (rest[1]! & 0x01) !== 0
    const bytes = count * (wide ? 2 : 1)
    if (2 + bytes > rest.length) return undefined
    const chars = rest.subarray(2, 2 + bytes)
    return { value: wide ? utf16.decode(chars) : compressed(chars), end: at + 2 + bytes }
  }
  if (rest.length < countBytes) return undefined
  const count = countBytes === 1 ? rest[0]! : rest[0]! | (rest[1]! << 8)
  if (count > text.maxLength) return null
  if (countBytes + count > rest.length) return undefined
  return { value: text.decoder.decode(rest.subarray(countBytes, countBytes + count)), end: at + countBytes + count }
}

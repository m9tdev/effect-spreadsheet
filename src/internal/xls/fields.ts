/**
 * The string forms BIFF8 and BIFF5 records use, and the text decoding of
 * the workbook's code page.
 *
 * @internal
 */
import { compressed, utf16 } from "../biff/segments.ts"
import { ok, type StringRead, tooLongString, TRUNCATED } from "../binary/layout.ts"

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

/*
 * Little-endian reads in a string's header, at offsets the caller checked
 * against the data's length (a string sits at a variable offset in its
 * record, so it checks its own bounds rather than a declared layout's).
 */
const u8At = (data: Uint8Array, at: number): number => data[at]!
const u16At = (data: Uint8Array, at: number): number => data[at]! | (data[at + 1]! << 8)
const u32At = (data: Uint8Array, at: number): number => (u16At(data, at) | (u16At(data, at + 2) << 16)) >>> 0

/** The characters of a BIFF8 string, 16-bit when `wide`, else 8-bit (UTF-16 low bytes). */
const chars8 = (bytes: Uint8Array, wide: boolean): string => (wide ? utf16.decode(bytes) : compressed(bytes))

/**
 * A BIFF8 XLUnicodeString (u16 count, flags, optional run count and
 * phonetic block size, characters, then the run and phonetic data) or
 * ShortXLUnicodeString (u8 count, flags, characters), or a BIFF5 string (u16
 * or u8 count, code-page bytes) at `at`, within one record's data. Returns
 * the text and the offset after it; `Truncated` if it runs past the data,
 * `TooLong` with its character count if that is over the limit. Each check
 * comes in the order the fields are read: a count over the limit is found
 * before the characters, and a cut header before the count is looked at.
 */
export const readString = (data: Uint8Array, at: number, text: Text, countBytes: 1 | 2): StringRead => {
  if (at > data.length) return TRUNCATED
  const left = data.length - at
  if (text.biff8) {
    if (countBytes === 2) {
      // cch u16, flags u8, then cRun u16 and cbExtRst u32 when the flags say so
      if (left < 3) return TRUNCATED
      const count = u16At(data, at)
      const flags = u8At(data, at + 2)
      const rich = (flags & 0x08) !== 0
      const ext = (flags & 0x04) !== 0
      const header = 3 + (rich ? 2 : 0) + (ext ? 4 : 0)
      if (left < header) return TRUNCATED
      const runs = rich ? u16At(data, at + 3) : 0
      const extSize = ext ? u32At(data, at + header - 4) : 0
      if (count > text.maxLength) return tooLongString(count)
      const wide = (flags & 0x01) !== 0
      const charsEnd = at + header + count * (wide ? 2 : 1)
      const end = charsEnd + runs * 4 + extSize
      if (end > data.length) return TRUNCATED
      return ok(chars8(data.subarray(at + header, charsEnd), wide), end)
    }
    // ShortXLUnicodeString: u8 count, flags, characters
    if (left < 2) return TRUNCATED
    const count = u8At(data, at)
    if (count > text.maxLength) return tooLongString(count)
    const wide = (u8At(data, at + 1) & 0x01) !== 0
    const end = at + 2 + count * (wide ? 2 : 1)
    if (end > data.length) return TRUNCATED
    return ok(chars8(data.subarray(at + 2, end), wide), end)
  }
  if (left < countBytes) return TRUNCATED
  const count = countBytes === 1 ? u8At(data, at) : u16At(data, at)
  if (count > text.maxLength) return tooLongString(count)
  const end = at + countBytes + count
  if (end > data.length) return TRUNCATED
  return ok(text.decoder.decode(data.subarray(at + countBytes, end)), end)
}

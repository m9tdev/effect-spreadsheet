/**
 * Typed reads of a start tag's attributes, one per kind of value: plain
 * text, OOXML text with `_xHHHH_` escapes, an unsigned integer, a boolean
 * and a relationship id. Each says what the attribute holds, or that it is
 * absent (or, for an integer, bad); what a missing or bad value means is
 * the caller's choice.
 *
 * Attribute names are matched as written (`name`, `r:id`).
 *
 * @internal
 */
import { attr, type XmlAttributes } from "./tokenizer.ts"

/** The value of attribute `name` as written, or `undefined` when it is absent. */
export const string = (attributes: XmlAttributes, name: string): string | undefined => attr(attributes, name)

/** `_xHHHH_` escapes in OOXML strings (ST_Xstring): one UTF-16 code unit each; `_x005F_` is `_`. */
export const unescapeXstring = (s: string): string =>
  s.indexOf("_x") < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)))

/** The value of attribute `name` with its `_xHHHH_` escapes decoded (ST_Xstring), or `undefined` when it is absent. */
export const xstring = (attributes: XmlAttributes, name: string): string | undefined => {
  const value = attr(attributes, name)
  return value === undefined ? undefined : unescapeXstring(value)
}

/** xsd:boolean's true (`1` or `true`); anything else, an absent attribute included, is false. */
export const bool = (attributes: XmlAttributes, name: string): boolean => {
  const value = attr(attributes, name)
  return value === "1" || value === "true"
}

/** How many decimal digits `max` has (`max` is a non-negative integer). */
const digitCount = (max: number): number => {
  let count = 1
  // progress: max shrinks tenfold each time, down to a single digit
  for (let rest = max; rest >= 10; rest = Math.floor(rest / 10)) count++
  return count
}

/**
 * `text` as an unsigned integer: ASCII digits only (no sign, space or
 * exponent), at most as many as `max` has (leading zeros count, so the
 * work is bounded by `max`, not by the text), and at most `max`. Else
 * `"bad"`.
 */
export const parseUint = (text: string, max: number): number | "bad" => {
  const length = text.length
  if (length === 0 || length > digitCount(max)) return "bad"
  let value = 0
  // progress: k increases up to length (at most the digits of max)
  for (let k = 0; k < length; k++) {
    const digit = text.charCodeAt(k) - 0x30
    if (digit < 0 || digit > 9) return "bad"
    value = value * 10 + digit
  }
  return value > max ? "bad" : value
}

/**
 * Attribute `name` as an unsigned integer up to `max` (see `parseUint`):
 * `undefined` when it is absent, `"bad"` when it isn't such a number.
 */
export const uint = (attributes: XmlAttributes, name: string, max: number): number | undefined | "bad" => {
  const value = attr(attributes, name)
  return value === undefined ? undefined : parseUint(value, max)
}

/**
 * A relationship attribute such as `r:id`: the first attribute whose local
 * name is `local` and that has a namespace prefix. The prefix isn't
 * resolved: any one is accepted (`r:id`, `x:id`), whatever namespace it is
 * bound to, or none. An unprefixed `id` doesn't count.
 */
export const rel = (attributes: XmlAttributes, local: "id"): string | undefined => {
  // progress: one attribute per iteration (bounded by maxXmlAttributes)
  for (const [key, value] of attributes) {
    const colon = key.indexOf(":")
    if (colon > 0 && key.slice(colon + 1) === local) return value
  }
  return undefined
}

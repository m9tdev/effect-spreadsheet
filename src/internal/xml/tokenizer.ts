/**
 * A small streaming XML tokenizer for spreadsheet parts.
 *
 * Feed it text in pieces of any size (a piece may end anywhere, even inside
 * a tag, an attribute value, an entity or a comment); it calls the handler
 * for start tags, end tags and text, in document order.
 *
 * Supported: elements and attributes (single or double quotes), text,
 * CDATA sections, comments and processing instructions (both skipped), the
 * five predefined entities and numeric character references. Not
 * supported, on purpose: DTDs (`<!DOCTYPE` fails with `MalformedFile`, so
 * there are no entity definitions to expand or fetch) and any other
 * declaration.
 *
 * Checked: one root element, matching end tags, nothing but whitespace
 * outside the root, unique attribute names, valid character references.
 * Limits: element depth (`maxXmlDepth`), attributes per element
 * (`maxXmlAttributes`), text, attribute values and the length of one tag,
 * comment or CDATA section (`maxStringLength`).
 *
 * Normalized as XML 1.0 says: line ends in text and CDATA (§2.11: CRLF and
 * a lone CR become LF), and literal whitespace in attribute values (§3.3.3:
 * tab, LF and CR become a space); characters written as references are
 * kept.
 *
 * Lenient on purpose (spreadsheet writers don't produce these, and
 * rejecting them buys no safety): control characters in text, `<` inside
 * attribute values, no space between attributes, a space after `</`, names
 * with more than one colon (the part after the first colon is the local
 * name), and an `<?xml ...?>` declaration anywhere (it is skipped like any
 * processing instruction).
 *
 * Errors carry the offset (in characters) of the tag or text run they are
 * about, the same however the input is split.
 *
 * Cost: every character is examined a bounded number of times (counted in
 * `work`, and handed to the work meter after each piece); a token split
 * across many pieces is collected in parts and joined once.
 *
 * @internal
 */
import { absurd } from "effect/Function"
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"
import { type Meter, noMeter } from "../meter.ts"

/** Attributes of a start tag: `[qualified name, decoded value]` pairs, in document order. */
export type XmlAttributes = ReadonlyArray<readonly [string, string]>

/** The value of attribute `name` (as written, e.g. `name`, `r:id`). */
export const attr = (attributes: XmlAttributes, name: string): string | undefined => {
  // progress: one attribute per iteration (bounded by maxXmlAttributes)
  for (const [key, value] of attributes) if (key === name) return value
  return undefined
}

/**
 * Receives tokens. Names are local names (any `prefix:` removed); attribute
 * names are kept as written (e.g. `r:id`). Consecutive `text` calls belong
 * to the same text run (text and CDATA arrive separately). A returned error
 * stops tokenizing.
 */
export interface XmlHandler {
  readonly open: (name: string, attributes: XmlAttributes) => SpreadsheetError | undefined | void
  readonly close: (name: string) => SpreadsheetError | undefined | void
  readonly text: (value: string) => SpreadsheetError | undefined | void
}

export interface XmlLimits {
  readonly maxXmlDepth: number
  readonly maxXmlAttributes: number
  readonly maxStringLength: number
}

/** Constructs that end with a multi-character terminator. */
type Delimiter = "Pi" | "Comment" | "Cdata"

/** Length of the opener ("?", "!--", "![CDATA[") that a terminator may not overlap. */
const OPENER: Record<Delimiter, number> = { Pi: 1, Comment: 3, Cdata: 8 }

const TERMINATOR: Record<Delimiter, string> = { Pi: "?>", Comment: "-->", Cdata: "]]>" }

/**
 * Markup (anything that begins with `<`) other than a start or end tag
 * complete within one piece, which the fast path handles directly. `parts`
 * collects everything after the `<` (`length` characters in all), across
 * pieces.
 * - `Undecided`: the first characters don't tell the construct yet (`<!-`
 *   could become a comment or a DTD).
 * - `Tag`: a start or end tag; `quote` is the quote character of the
 *   attribute value a start tag is inside of, or 0.
 * - `Delimited`: a processing instruction, comment or CDATA section; `tail`
 *   holds its last characters past the opener, for a terminator split
 *   across pieces.
 */
type Markup = Undecided | Tag | Delimited
type Undecided = { readonly _tag: "Undecided"; readonly parts: Array<string>; length: number }
type Tag = { readonly _tag: "Tag"; readonly start: boolean; readonly parts: Array<string>; length: number; quote: number }
type Delimited = {
  readonly _tag: "Delimited"
  readonly kind: Delimiter
  readonly parts: Array<string>
  length: number
  tail: string
}

/** Where the tokenizer is between pieces: in text, or inside markup. */
type Mode = { readonly _tag: "Text" } | Markup

/** Shared, so that text costs no allocation. */
const TEXT: Mode = { _tag: "Text" }

/**
 * What a step of the scan returns: undefined to go on, or the error that
 * stops tokenizing. A handler's result is one already.
 */
type Status = SpreadsheetError | undefined | void

/** What a scan over markup returns: the position to go on from, or the error that stops tokenizing. */
type Scanned = number | SpreadsheetError

/** Prolog: before the root element; Root: inside it; Epilog: after it closed. */
type Phase = "Prolog" | "Root" | "Epilog"

/** Extra room for markup beyond `maxStringLength` (tag name, attribute names, quotes). */
const MARKUP_SLACK = 64 * 1024

/** XML 1.0 §2.11: CRLF and a lone CR become LF. */
const normalizeLineEnds = (s: string): string => (s.indexOf("\r") < 0 ? s : s.replace(/\r\n?/g, "\n"))

/** XML 1.0 §3.3.3 (CDATA-type attributes): after §2.11, each literal tab, LF or CR becomes a space. */
const normalizeAttribute = (s: string): string => (/[\t\n\r]/.test(s) ? s.replace(/\r\n|[\t\n\r]/g, " ") : s)

const localName = (qname: string): string => {
  const colon = qname.indexOf(":")
  return colon < 0 ? qname : qname.slice(colon + 1)
}

const isSpace = (c: number): boolean => c === 0x20 || c === 0x09 || c === 0x0A || c === 0x0D

/**
 * A rough XML name check: non-empty, no whitespace or markup characters
 * (`< > / = ! ? " ' &`), not starting with `-`, `.` or a digit.
 * "Whitespace" is exactly JavaScript's `\s` (the earlier regex used it):
 * tab to CR, space, U+00A0, U+1680, U+2000-U+200A, U+2028, U+2029,
 * U+202F, U+205F, U+3000, U+FEFF. Other control characters are allowed.
 */
export const isName = (s: string): boolean => {
  const n = s.length
  if (n === 0) return false
  const first = s.charCodeAt(0)
  if (first === 0x2D || first === 0x2E || (first >= 0x30 && first <= 0x39)) return false
  // progress: k increases up to n
  for (let k = 0; k < n; k++) {
    const c = s.charCodeAt(k)
    if (
      (c >= 0x09 && c <= 0x0D) || c === 0x20 || c === 0x3C || c === 0x3E || c === 0x2F || c === 0x3D || c === 0x21 ||
      c === 0x3F || c === 0x22 || c === 0x27 || c === 0x26 || c === 0xA0 || c === 0x1680 ||
      (c >= 0x2000 && c <= 0x200A) || c === 0x2028 || c === 0x2029 || c === 0x202F || c === 0x205F ||
      c === 0x3000 || c === 0xFEFF
    ) return false
  }
  return true
}

export class XmlTokenizer {
  /** Characters examined in the current piece (each input character a bounded number of times). */
  private work = 0

  private readonly meter: Meter
  private readonly handler: XmlHandler
  private readonly limits: XmlLimits
  private readonly part: string | undefined

  private mode: Mode = TEXT

  private text: Array<string> = []
  private textLength = 0

  private readonly stack: Array<string> = []
  private phase: Phase = "Prolog"
  /** Characters fed before the current piece. */
  private consumed = 0
  /** Offset of the current markup's '<' (for errors). */
  private markupStart = 0
  /** Offset of the current text run's first character (for errors). */
  private textStart = 0
  /** The error that stopped tokenizing: every later `feed` and `end` returns it. */
  private failure: SpreadsheetError | undefined

  constructor(handler: XmlHandler, limits: XmlLimits, part?: string, meter: Meter = noMeter) {
    this.handler = handler
    this.limits = limits
    this.part = part
    this.meter = meter
  }

  /** Hands the piece's work to the meter. */
  private report(): void {
    if (this.work > 0) {
      this.meter.add("characters", this.work)
      this.work = 0
    }
  }

  private malformed(reason: string, offset?: number): MalformedFile {
    return new MalformedFile({
      format: "xml",
      reason,
      ...(this.part === undefined ? {} : { part: this.part }),
      ...(offset === undefined ? {} : { offset })
    })
  }

  /** Feeds the next piece of text. Returns the error that stopped tokenizing, if any. */
  feed(chunk: string): SpreadsheetError | undefined {
    if (this.failure !== undefined) return this.failure
    const n = chunk.length
    let error: Status
    let i = 0
    // progress: every branch moves i forward by at least 1, or stops the
    // loop with an error
    while (i < n) {
      const mode = this.mode
      if (mode._tag === "Text") {
        const lt = chunk.indexOf("<", i)
        const end = lt < 0 ? n : lt
        this.work += end - i + 1
        if (end > i) {
          if (this.textLength === 0) this.textStart = this.consumed + i
          error = this.addText(chunk.slice(i, end))
          if (error !== undefined) break
        }
        if (lt < 0) {
          i = n
        } else {
          error = this.flushText()
          if (error !== undefined) break
          this.markupStart = this.consumed + lt
          const whole = this.wholeTag(chunk, lt)
          if (typeof whole !== "number") {
            error = whole
            break
          }
          if (whole > lt) {
            i = whole // the tag was complete in this piece and is handled
          } else {
            this.mode = { _tag: "Undecided", parts: [], length: 0 }
            i = lt + 1
          }
        }
      } else {
        const next = this.scanMarkup(mode, chunk, i)
        if (typeof next !== "number") {
          error = next
          break
        }
        i = next
      }
    }
    this.consumed += n
    this.report()
    if (error === undefined) return undefined
    return (this.failure = error)
  }

  /** Ends the input. Returns the error, if the document is incomplete or invalid. */
  end(): SpreadsheetError | undefined {
    if (this.failure !== undefined) return this.failure
    if (this.mode._tag !== "Text") return (this.failure = this.malformed("truncated inside a tag", this.consumed))
    const error = this.flushText()
    this.report()
    if (error !== undefined) return (this.failure = error)
    if (this.stack.length > 0) {
      return (this.failure = this.malformed(`truncated: ${this.stack.length} element(s) not closed`, this.consumed))
    }
    if (this.phase === "Prolog") return (this.failure = this.malformed("no root element", this.consumed))
    return undefined
  }

  private addText(s: string): Status {
    this.text.push(s)
    this.textLength += s.length
    if (this.textLength > this.limits.maxStringLength) {
      return new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: this.textLength })
    }
    return undefined
  }

  private flushText(): Status {
    if (this.textLength === 0) return undefined
    const raw = this.text.length === 1 ? this.text[0]! : this.text.join("")
    this.text = []
    this.textLength = 0
    this.work += raw.length
    if (this.stack.length === 0) {
      // only whitespace may sit outside the root element
      // progress: k increases up to raw.length
      for (let k = 0; k < raw.length; k++) {
        if (!isSpace(raw.charCodeAt(k))) return this.malformed("text outside the root element", this.textStart)
      }
      return undefined
    }
    // line ends are normalized first (XML 1.0 §2.11), so &#13; survives
    const value = this.decode(normalizeLineEnds(raw), this.textStart)
    if (typeof value !== "string") return value
    return this.handler.text(value)
  }

  private addMarkup(mode: Markup, s: string): Status {
    mode.parts.push(s)
    mode.length += s.length
    if (mode.length > this.limits.maxStringLength + MARKUP_SLACK) {
      return new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: mode.length })
    }
    return undefined
  }

  /**
   * Fast path for the common case: a start or end tag that ends within this
   * piece. Handles it exactly as the general path would (same markup, same
   * checks, same work count) and returns the position after its `>`; or -1
   * (nothing done) when the tag is something else or continues in the next
   * piece; or the error that stops tokenizing.
   */
  private wholeTag(chunk: string, lt: number): Scanned {
    const n = chunk.length
    const first = chunk.charCodeAt(lt + 1)
    if (lt + 1 >= n || first === 0x21 || first === 0x3F) return -1
    const start = first !== 0x2F
    let quote = 0
    let j = lt + 1
    // progress: j increases up to n
    for (; j < n; j++) {
      const c = chunk.charCodeAt(j)
      if (quote !== 0) {
        if (c === quote) quote = 0
      } else if (c === 0x3E) break
      else if (start && (c === 0x22 || c === 0x27)) quote = c
    }
    if (j >= n) return -1
    const length = j - lt - 1
    this.work += 2 * length + 1
    if (length > this.limits.maxStringLength + MARKUP_SLACK) {
      return new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: length })
    }
    const markup = chunk.slice(lt + 1, j)
    const error = start ? this.startTag(markup) : this.endTag(markup)
    return error ?? j + 1
  }

  /** Scans markup from `chunk[i]`; returns the next position (> i), or the error that stops tokenizing. */
  private scanMarkup(mode: Markup, chunk: string, i: number): Scanned {
    switch (mode._tag) {
      case "Undecided":
        return this.scanUndecided(mode, chunk, i)
      case "Tag":
        return this.scanTag(mode, chunk, i)
      case "Delimited":
        return this.scanDelimited(mode, chunk, i)
      default:
        return absurd(mode)
    }
  }

  /**
   * Decides the construct from its first characters; they may arrive one
   * piece at a time, so look at what we have so far.
   */
  private scanUndecided(mode: Undecided, chunk: string, i: number): Scanned {
    const n = chunk.length
    const head = (mode.parts.join("") + chunk.slice(i, i + 9))
    const c = head.charCodeAt(0)
    if (c === 0x2F) return this.scanTag(this.decide(mode, "End"), chunk, i)
    if (c === 0x3F) return this.scanDelimited(this.decide(mode, "Pi"), chunk, i)
    if (c !== 0x21) return this.scanTag(this.decide(mode, "Start"), chunk, i)
    if (head.startsWith("!--")) return this.scanDelimited(this.decide(mode, "Comment"), chunk, i)
    if (head.startsWith("![CDATA[")) return this.scanDelimited(this.decide(mode, "Cdata"), chunk, i)
    if (head.startsWith("!DOCTYPE")) return this.malformed("DTDs (<!DOCTYPE>) are not supported", this.markupStart)
    if (!("!--".startsWith(head) || "![CDATA[".startsWith(head) || "!DOCTYPE".startsWith(head))) {
      return this.malformed("unsupported declaration", this.markupStart)
    }
    // not decidable yet: take the rest of this piece (it's short) and wait
    const take = Math.min(n, i + Math.max(1, head.length + 1 - mode.length))
    this.work += take - i
    return this.addMarkup(mode, chunk.slice(i, take)) ?? take
  }

  /** The transition out of `Undecided`: the decided mode takes over the characters collected so far. */
  private decide(mode: Undecided, kind: "Start" | "End"): Tag
  private decide(mode: Undecided, kind: Delimiter): Delimited
  private decide(mode: Undecided, kind: "Start" | "End" | Delimiter): Tag | Delimited {
    const next: Tag | Delimited = kind === "Start" || kind === "End"
      ? { _tag: "Tag", start: kind === "Start", parts: mode.parts, length: mode.length, quote: 0 }
      : { _tag: "Delimited", kind, parts: mode.parts, length: mode.length, tail: "" }
    this.mode = next
    return next
  }

  /** A start or end tag: finds '>' outside quotes (quotes only matter in start tags). */
  private scanTag(mode: Tag, chunk: string, i: number): Scanned {
    const n = chunk.length
    const start = mode.start
    let quote = mode.quote
    let j = i
    // progress: j increases up to n
    for (; j < n; j++) {
      const c = chunk.charCodeAt(j)
      if (quote !== 0) {
        if (c === quote) quote = 0
      } else if (c === 0x3E) break
      else if (start && (c === 0x22 || c === 0x27)) quote = c
    }
    mode.quote = quote
    this.work += j - i + 1
    const error = this.addMarkup(mode, chunk.slice(i, j))
    if (error !== undefined) return error
    if (j === n) return n
    return this.finishMarkup(mode) ?? j + 1
  }

  /**
   * PI, comment, CDATA: a multi-character terminator, possibly split across
   * pieces. It may not reuse characters of the opener ("<!-->" is not a
   * comment), so only markup positions past the opener count, and the same
   * holds whatever the splitting.
   */
  private scanDelimited(mode: Delimited, chunk: string, i: number): Scanned {
    const n = chunk.length
    const terminator = TERMINATOR[mode.kind]
    const opener = OPENER[mode.kind]
    let endInChunk = -1
    if (mode.tail.length > 0) {
      // a terminator that starts in the tail ends within the next few characters
      const joint = mode.tail + chunk.slice(i, i + terminator.length - 1)
      const k = joint.indexOf(terminator)
      this.work += joint.length
      if (k >= 0) endInChunk = i + k + terminator.length - mode.tail.length
    }
    if (endInChunk < 0) {
      const from = i + Math.max(0, opener - mode.length)
      const at = from <= n ? chunk.indexOf(terminator, from) : -1
      if (at < 0) {
        this.work += n - i
        const error = this.addMarkup(mode, chunk.slice(i))
        if (error !== undefined) return error
        // keep the last characters past the opener, for a terminator split
        // across pieces
        const keep = Math.min(terminator.length - 1, Math.max(0, mode.length - opener))
        const last = (mode.tail + chunk.slice(Math.max(i, n - keep))).slice(-keep)
        mode.tail = keep === 0 ? "" : last
        return n
      }
      endInChunk = at + terminator.length
      this.work += endInChunk - i
    }
    return this.addMarkup(mode, chunk.slice(i, endInChunk)) ?? this.finishMarkup(mode) ?? endInChunk
  }

  /** The transition back to `Text`: handles the completed construct. */
  private finishMarkup(mode: Tag | Delimited): Status {
    this.mode = TEXT
    const markup = mode.parts.length === 1 ? mode.parts[0]! : mode.parts.join("")
    this.work += markup.length
    switch (mode._tag) {
      case "Tag":
        return mode.start ? this.startTag(markup) : this.endTag(markup)
      case "Delimited":
        return this.finishDelimited(mode.kind, markup)
      default:
        return absurd(mode)
    }
  }

  private finishDelimited(kind: Delimiter, markup: string): Status {
    switch (kind) {
      case "Pi":
      case "Comment":
        return undefined
      case "Cdata": {
        if (this.stack.length === 0) return this.malformed("CDATA outside the root element", this.markupStart)
        // markup is "![CDATA[" + content + "]]>"
        const content = markup.slice(8, markup.length - 3)
        if (content.length > this.limits.maxStringLength) {
          return new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: content.length })
        }
        return content.length > 0 ? this.handler.text(normalizeLineEnds(content)) : undefined
      }
      default:
        return absurd(kind)
    }
  }

  private startTag(markup: string): Status {
    // markup is everything between '<' and '>'
    let body = markup
    let selfClosing = false
    if (body.charCodeAt(body.length - 1) === 0x2F) {
      selfClosing = true
      body = body.slice(0, -1)
    }
    let p = 0
    const n = body.length
    // progress: p increases up to the end of the name
    while (p < n && !isSpace(body.charCodeAt(p))) p++
    const qname = body.slice(0, p)
    if (!isName(qname)) return this.malformed(`bad element name ${JSON.stringify(qname.slice(0, 40))}`, this.markupStart)
    if (this.phase === "Epilog") return this.malformed("more than one root element", this.markupStart)
    const attributes: Array<readonly [string, string]> = []
    // duplicates: a linear check for the usual few attributes, a Set past that
    let seen: Set<string> | undefined
    // progress: each iteration consumes one attribute (p moves past its
    // closing quote) or ends the loop
    while (p < n) {
      while (p < n && isSpace(body.charCodeAt(p))) p++
      if (p >= n) break
      const eq = body.indexOf("=", p)
      if (eq < 0) return this.malformed(`attribute without a value in <${qname}>`, this.markupStart)
      const name = body.slice(p, eq).trim()
      let q = eq + 1
      while (q < n && isSpace(body.charCodeAt(q))) q++
      const quote = body.charCodeAt(q)
      if (quote !== 0x22 && quote !== 0x27) return this.malformed(`unquoted attribute value in <${qname}>`, this.markupStart)
      const close = body.indexOf(String.fromCharCode(quote), q + 1)
      if (close < 0 || !isName(name)) return this.malformed(`bad attribute in <${qname}>`, this.markupStart)
      let duplicate = false
      if (seen !== undefined) duplicate = seen.has(name)
      else {
        // progress: k increases up to attributes.length (< 16 here)
        for (let k = 0; k < attributes.length && !duplicate; k++) duplicate = attributes[k]![0] === name
      }
      if (duplicate) return this.malformed(`duplicate attribute ${name} in <${qname}>`, this.markupStart)
      if (seen !== undefined) seen.add(name)
      else if (attributes.length >= 15) {
        seen = new Set(attributes.map((a) => a[0]))
        seen.add(name)
      }
      if (attributes.length >= this.limits.maxXmlAttributes) {
        return new LimitExceeded({
          limit: "maxXmlAttributes",
          max: this.limits.maxXmlAttributes,
          actual: attributes.length + 1
        })
      }
      const raw = body.slice(q + 1, close)
      if (raw.length > this.limits.maxStringLength) {
        return new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: raw.length })
      }
      // literal whitespace becomes a space (XML 1.0 §3.3.3) before references
      // are decoded, so &#10; and &#9; survive
      const value = this.decode(normalizeAttribute(raw), this.markupStart)
      if (typeof value !== "string") return value
      attributes.push([name, value])
      p = close + 1
    }
    if (this.stack.length >= this.limits.maxXmlDepth) {
      return new LimitExceeded({ limit: "maxXmlDepth", max: this.limits.maxXmlDepth, actual: this.stack.length + 1 })
    }
    this.phase = "Root"
    this.stack.push(qname)
    const error = this.handler.open(localName(qname), attributes)
    if (error !== undefined || !selfClosing) return error
    return this.closeElement(qname)
  }

  private endTag(markup: string): Status {
    const qname = markup.slice(1).trim()
    const top = this.stack[this.stack.length - 1]
    if (top === undefined) return this.malformed(`end tag </${qname}> without a start tag`, this.markupStart)
    if (top !== qname) return this.malformed(`end tag </${qname}> does not match <${top}>`, this.markupStart)
    return this.closeElement(qname)
  }

  private closeElement(qname: string): Status {
    this.stack.pop()
    if (this.stack.length === 0) this.phase = "Epilog"
    return this.handler.close(localName(qname))
  }

  /** Decodes entities and character references; or the error, if one is invalid. */
  private decode(raw: string, offset: number): string | MalformedFile {
    let amp = raw.indexOf("&")
    if (amp < 0) return raw
    const parts: Array<string> = []
    let from = 0
    // progress: amp moves past each reference (to the next '&' after its ';')
    while (amp >= 0) {
      // (linear: on success the scan continues after ';', on failure it stops)
      const semi = raw.indexOf(";", amp + 1)
      if (semi < 0) return this.malformed("bad entity reference", offset)
      const name = raw.slice(amp + 1, semi)
      let ch: string | undefined
      if (name === "lt") ch = "<"
      else if (name === "gt") ch = ">"
      else if (name === "amp") ch = "&"
      else if (name === "quot") ch = "\""
      else if (name === "apos") ch = "'"
      else if (name.charCodeAt(0) === 0x23) {
        const hex = name.charCodeAt(1) === 0x78 || name.charCodeAt(1) === 0x58
        // leading zeros are allowed in any number (&#00000065; is "A")
        const digits = name.slice(hex ? 2 : 1).replace(/^0+(?=.)/, "")
        const valid = hex ? /^[0-9A-Fa-f]{1,6}$/.test(digits) : /^[0-9]{1,7}$/.test(digits)
        const code = valid ? Number.parseInt(digits, hex ? 16 : 10) : -1
        if (
          code === 0x9 || code === 0xA || code === 0xD || (code >= 0x20 && code <= 0xD7FF) ||
          (code >= 0xE000 && code <= 0xFFFD) || (code >= 0x10000 && code <= 0x10FFFF)
        ) ch = String.fromCodePoint(code)
      }
      if (ch === undefined) return this.malformed(`unknown or invalid entity &${name.slice(0, 12)};`, offset)
      parts.push(raw.slice(from, amp), ch)
      from = semi + 1
      amp = raw.indexOf("&", from)
    }
    parts.push(raw.slice(from))
    return parts.join("")
  }
}

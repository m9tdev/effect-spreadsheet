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
 * Cost: every character is examined a bounded number of times (the
 * `work` counter); a token split across many pieces is collected in parts
 * and joined once.
 *
 * @internal
 */
import { LimitExceeded, MalformedFile, type SpreadsheetError } from "../../SpreadsheetError.ts"

/** Attributes of a start tag: `[qualified name, decoded value]` pairs, in document order. */
export type XmlAttributes = ReadonlyArray<readonly [string, string]>

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

const Kind = {
  Unknown: 0,
  Start: 1,
  End: 2,
  Pi: 3,
  Comment: 4,
  Cdata: 5
} as const
type Kind = typeof Kind[keyof typeof Kind]

/** Length of the opener ("?", "!--", "![CDATA[") that a terminator may not overlap. */
const OPENER: Record<number, number> = {
  [Kind.Pi]: 1,
  [Kind.Comment]: 3,
  [Kind.Cdata]: 8
}

const TERMINATOR: Record<number, string> = {
  [Kind.Pi]: "?>",
  [Kind.Comment]: "-->",
  [Kind.Cdata]: "]]>"
}

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
  /** Characters examined so far (each input character a bounded number of times). */
  work = 0

  private readonly handler: XmlHandler
  private readonly limits: XmlLimits
  private readonly part: string | undefined

  private inMarkup = false
  private kind: Kind = Kind.Unknown
  private markup: Array<string> = []
  private markupLength = 0
  /** In a start tag: the quote character we're inside of, or 0. */
  private quote = 0
  /** For multi-character terminators: the last characters of the markup so far. */
  private tail = ""

  private text: Array<string> = []
  private textLength = 0

  private readonly stack: Array<string> = []
  private rootSeen = false
  private rootClosed = false
  /** Characters fed before the current piece. */
  private consumed = 0
  /** Offset of the current markup's '<' (for errors). */
  private markupStart = 0
  /** Offset of the current text run's first character (for errors). */
  private textStart = 0
  private failure: SpreadsheetError | undefined

  constructor(handler: XmlHandler, limits: XmlLimits, part?: string) {
    this.handler = handler
    this.limits = limits
    this.part = part
  }

  private malformed(reason: string, offset?: number): MalformedFile {
    return new MalformedFile({
      format: "xml",
      reason,
      ...(this.part === undefined ? {} : { part: this.part }),
      ...(offset === undefined ? {} : { offset })
    })
  }

  private fail(error: SpreadsheetError): void {
    if (this.failure === undefined) this.failure = error
  }

  /** Feeds the next piece of text. Returns the error that stopped tokenizing, if any. */
  feed(chunk: string): SpreadsheetError | undefined {
    if (this.failure !== undefined) return this.failure
    const n = chunk.length
    let i = 0
    // progress: every branch moves i forward by at least 1, or fails
    while (i < n && this.failure === undefined) {
      if (!this.inMarkup) {
        const lt = chunk.indexOf("<", i)
        const end = lt < 0 ? n : lt
        this.work += end - i + 1
        if (end > i) {
          if (this.textLength === 0) this.textStart = this.consumed + i
          this.addText(chunk.slice(i, end))
        }
        if (lt < 0) {
          i = n
        } else {
          this.flushText()
          this.markupStart = this.consumed + lt
          const whole = this.failure === undefined ? this.wholeTag(chunk, lt) : -1
          if (whole > lt) {
            i = whole // the tag was complete in this piece and is handled
          } else {
            this.inMarkup = true
            this.kind = Kind.Unknown
            this.markup = []
            this.markupLength = 0
            this.quote = 0
            this.tail = ""
            i = lt + 1
          }
        }
      } else {
        i = this.scanMarkup(chunk, i)
      }
    }
    this.consumed += n
    return this.failure
  }

  /** Ends the input. Returns the error, if the document is incomplete or invalid. */
  end(): SpreadsheetError | undefined {
    if (this.failure !== undefined) return this.failure
    if (this.inMarkup) return (this.failure = this.malformed("truncated inside a tag", this.consumed))
    this.flushText()
    if (this.failure !== undefined) return this.failure
    if (this.stack.length > 0) {
      return (this.failure = this.malformed(`truncated: ${this.stack.length} element(s) not closed`, this.consumed))
    }
    if (!this.rootSeen) return (this.failure = this.malformed("no root element", this.consumed))
    return undefined
  }

  private addText(s: string): void {
    this.text.push(s)
    this.textLength += s.length
    if (this.textLength > this.limits.maxStringLength) {
      this.fail(new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: this.textLength }))
    }
  }

  private flushText(): void {
    if (this.textLength === 0) return
    const raw = this.text.length === 1 ? this.text[0]! : this.text.join("")
    this.text = []
    this.textLength = 0
    this.work += raw.length
    if (this.stack.length === 0) {
      // only whitespace may sit outside the root element
      // progress: k increases up to raw.length
      for (let k = 0; k < raw.length; k++) {
        if (!isSpace(raw.charCodeAt(k))) {
          this.fail(this.malformed("text outside the root element", this.textStart))
          return
        }
      }
      return
    }
    // line ends are normalized first (XML 1.0 §2.11), so &#13; survives
    const value = this.decode(normalizeLineEnds(raw), this.textStart)
    if (value === undefined) return
    this.emit(this.handler.text(value))
  }

  private emit(result: SpreadsheetError | undefined | void): void {
    if (result !== undefined) this.fail(result)
  }

  private addMarkup(s: string): void {
    this.markup.push(s)
    this.markupLength += s.length
    if (this.markupLength > this.limits.maxStringLength + MARKUP_SLACK) {
      this.fail(new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: this.markupLength }))
    }
  }

  /**
   * Fast path for the common case: a start or end tag that ends within this
   * piece. Handles it exactly as the general path would (same markup, same
   * checks, same work count) and returns the position after its `>`; or -1
   * (nothing done) when the tag is something else or continues in the next
   * piece.
   */
  private wholeTag(chunk: string, lt: number): number {
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
      this.fail(new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: length }))
      return j + 1
    }
    const markup = chunk.slice(lt + 1, j)
    this.kind = start ? Kind.Start : Kind.End
    if (start) this.startTag(markup)
    else this.endTag(markup)
    return j + 1
  }

  /** Scans markup from `chunk[i]`; returns the next position (> i, or i on failure). */
  private scanMarkup(chunk: string, i: number): number {
    const n = chunk.length
    if (this.kind === Kind.Unknown) {
      // Decide the construct from its first characters; they may arrive
      // one piece at a time, so look at what we have so far.
      const head = (this.markup.join("") + chunk.slice(i, i + 9))
      const c = head.charCodeAt(0)
      let kind: Kind = Kind.Unknown
      let need = 1
      if (c === 0x2F) kind = Kind.End
      else if (c === 0x3F) kind = Kind.Pi
      else if (c === 0x21) {
        if (head.startsWith("!--")) kind = Kind.Comment
        else if (head.startsWith("![CDATA[")) kind = Kind.Cdata
        else if ("!--".startsWith(head) || "![CDATA[".startsWith(head)) need = head.length + 1
        else {
          this.fail(
            this.malformed(
              head.startsWith("!DOCTYPE") || "!DOCTYPE".startsWith(head)
                ? "DTDs (<!DOCTYPE>) are not supported"
                : "unsupported declaration",
              this.markupStart
            )
          )
          return i + 1
        }
      } else kind = Kind.Start
      if (kind === Kind.Unknown) {
        // not decidable yet: take the rest of this piece (it's short) and wait
        const take = Math.min(n, i + Math.max(1, need - this.markupLength))
        this.work += take - i
        this.addMarkup(chunk.slice(i, take))
        return take
      }
      this.kind = kind
    }
    if (this.kind === Kind.Start || this.kind === Kind.End) {
      // find '>' outside quotes (quotes only matter in start tags)
      let j = i
      const start = this.kind === Kind.Start
      // progress: j increases up to n
      for (; j < n; j++) {
        const c = chunk.charCodeAt(j)
        if (this.quote !== 0) {
          if (c === this.quote) this.quote = 0
        } else if (c === 0x3E) break
        else if (start && (c === 0x22 || c === 0x27)) this.quote = c
      }
      this.work += j - i + 1
      this.addMarkup(chunk.slice(i, j))
      if (j === n) return n
      this.finishMarkup()
      return j + 1
    }
    // PI, comment, CDATA: a multi-character terminator, possibly split
    // across pieces. It may not reuse characters of the opener ("<!-->" is
    // not a comment), so only markup positions past the opener count, and
    // the same holds whatever the splitting.
    const terminator = TERMINATOR[this.kind]!
    const opener = OPENER[this.kind]!
    let endInChunk = -1
    if (this.tail.length > 0) {
      // a terminator that starts in the tail ends within the next few characters
      const joint = this.tail + chunk.slice(i, i + terminator.length - 1)
      const k = joint.indexOf(terminator)
      this.work += joint.length
      if (k >= 0) endInChunk = i + k + terminator.length - this.tail.length
    }
    if (endInChunk < 0) {
      const from = i + Math.max(0, opener - this.markupLength)
      const at = from <= n ? chunk.indexOf(terminator, from) : -1
      if (at < 0) {
        this.work += n - i
        this.addMarkup(chunk.slice(i))
        // keep the last characters past the opener, for a terminator split
        // across pieces
        const keep = Math.min(terminator.length - 1, Math.max(0, this.markupLength - opener))
        const last = (this.tail + chunk.slice(Math.max(i, n - keep))).slice(-keep)
        this.tail = keep === 0 ? "" : last
        return n
      }
      endInChunk = at + terminator.length
      this.work += endInChunk - i
    }
    this.addMarkup(chunk.slice(i, endInChunk))
    this.finishMarkup()
    return endInChunk
  }

  private finishMarkup(): void {
    if (this.failure !== undefined) return
    const markup = this.markup.length === 1 ? this.markup[0]! : this.markup.join("")
    this.markup = []
    this.markupLength = 0
    this.inMarkup = false
    this.tail = ""
    this.work += markup.length
    switch (this.kind) {
      case Kind.Start:
        return this.startTag(markup)
      case Kind.End:
        return this.endTag(markup)
      case Kind.Pi:
      case Kind.Comment:
        return
      case Kind.Cdata: {
        if (this.stack.length === 0) {
          this.fail(this.malformed("CDATA outside the root element", this.markupStart))
          return
        }
        // markup is "![CDATA[" + content + "]]>"
        const content = markup.slice(8, markup.length - 3)
        if (content.length > this.limits.maxStringLength) {
          this.fail(new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: content.length }))
          return
        }
        if (content.length > 0) this.emit(this.handler.text(normalizeLineEnds(content)))
        return
      }
    }
  }

  private startTag(markup: string): void {
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
    if (!isName(qname)) {
      this.fail(this.malformed(`bad element name ${JSON.stringify(qname.slice(0, 40))}`, this.markupStart))
      return
    }
    if (this.rootClosed) {
      this.fail(this.malformed("more than one root element", this.markupStart))
      return
    }
    const attributes: Array<readonly [string, string]> = []
    // duplicates: a linear check for the usual few attributes, a Set past that
    let seen: Set<string> | undefined
    // progress: each iteration consumes one attribute (p moves past its
    // closing quote) or ends the loop
    while (p < n) {
      while (p < n && isSpace(body.charCodeAt(p))) p++
      if (p >= n) break
      const eq = body.indexOf("=", p)
      if (eq < 0) {
        this.fail(this.malformed(`attribute without a value in <${qname}>`, this.markupStart))
        return
      }
      const name = body.slice(p, eq).trim()
      let q = eq + 1
      while (q < n && isSpace(body.charCodeAt(q))) q++
      const quote = body.charCodeAt(q)
      if (quote !== 0x22 && quote !== 0x27) {
        this.fail(this.malformed(`unquoted attribute value in <${qname}>`, this.markupStart))
        return
      }
      const close = body.indexOf(String.fromCharCode(quote), q + 1)
      if (close < 0 || !isName(name)) {
        this.fail(this.malformed(`bad attribute in <${qname}>`, this.markupStart))
        return
      }
      let duplicate = false
      if (seen !== undefined) duplicate = seen.has(name)
      else {
        // progress: k increases up to attributes.length (< 16 here)
        for (let k = 0; k < attributes.length && !duplicate; k++) duplicate = attributes[k]![0] === name
      }
      if (duplicate) {
        this.fail(this.malformed(`duplicate attribute ${name} in <${qname}>`, this.markupStart))
        return
      }
      if (seen !== undefined) seen.add(name)
      else if (attributes.length >= 15) {
        seen = new Set(attributes.map((a) => a[0]))
        seen.add(name)
      }
      if (attributes.length >= this.limits.maxXmlAttributes) {
        this.fail(
          new LimitExceeded({ limit: "maxXmlAttributes", max: this.limits.maxXmlAttributes, actual: attributes.length + 1 })
        )
        return
      }
      const raw = body.slice(q + 1, close)
      if (raw.length > this.limits.maxStringLength) {
        this.fail(new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: raw.length }))
        return
      }
      // literal whitespace becomes a space (XML 1.0 §3.3.3) before references
      // are decoded, so &#10; and &#9; survive
      const value = this.decode(normalizeAttribute(raw), this.markupStart)
      if (value === undefined) return
      attributes.push([name, value])
      p = close + 1
    }
    if (this.stack.length >= this.limits.maxXmlDepth) {
      this.fail(new LimitExceeded({ limit: "maxXmlDepth", max: this.limits.maxXmlDepth, actual: this.stack.length + 1 }))
      return
    }
    this.rootSeen = true
    this.stack.push(qname)
    const name = localName(qname)
    this.emit(this.handler.open(name, attributes))
    if (selfClosing && this.failure === undefined) this.closeElement(qname)
  }

  private endTag(markup: string): void {
    const qname = markup.slice(1).trim()
    const top = this.stack[this.stack.length - 1]
    if (top === undefined) {
      this.fail(this.malformed(`end tag </${qname}> without a start tag`, this.markupStart))
      return
    }
    if (top !== qname) {
      this.fail(this.malformed(`end tag </${qname}> does not match <${top}>`, this.markupStart))
      return
    }
    this.closeElement(qname)
  }

  private closeElement(qname: string): void {
    this.stack.pop()
    if (this.stack.length === 0) this.rootClosed = true
    this.emit(this.handler.close(localName(qname)))
  }

  /** Decodes entities and character references; undefined (and a failure) if one is invalid. */
  private decode(raw: string, offset: number): string | undefined {
    let amp = raw.indexOf("&")
    if (amp < 0) return raw
    const parts: Array<string> = []
    let from = 0
    // progress: amp moves past each reference (to the next '&' after its ';')
    while (amp >= 0) {
      // (linear: on success the scan continues after ';', on failure it stops)
      const semi = raw.indexOf(";", amp + 1)
      if (semi < 0) {
        this.fail(this.malformed("bad entity reference", offset))
        return undefined
      }
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
      if (ch === undefined) {
        this.fail(this.malformed(`unknown or invalid entity &${name.slice(0, 12)};`, offset))
        return undefined
      }
      parts.push(raw.slice(from, amp), ch)
      from = semi + 1
      amp = raw.indexOf("&", from)
    }
    parts.push(raw.slice(from))
    return parts.join("")
  }
}

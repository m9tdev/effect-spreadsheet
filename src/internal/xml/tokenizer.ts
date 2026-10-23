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

const TERMINATOR: Record<number, string> = {
  [Kind.Pi]: "?>",
  [Kind.Comment]: "-->",
  [Kind.Cdata]: "]]>"
}

/** Extra room for markup beyond `maxStringLength` (tag name, attribute names, quotes). */
const MARKUP_SLACK = 64 * 1024

const localName = (qname: string): string => {
  const colon = qname.indexOf(":")
  return colon < 0 ? qname : qname.slice(colon + 1)
}

const isSpace = (c: number): boolean => c === 0x20 || c === 0x09 || c === 0x0A || c === 0x0D

/** A rough XML name check: no whitespace or markup characters, not starting with `-`, `.` or a digit. */
const NAME = /^[^\s<>/=!?"'&\-.\d][^\s<>/=!?"'&]*$/

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
  /** Characters fed before the current piece (for error offsets). */
  private consumed = 0
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
        if (end > i) this.addText(chunk.slice(i, end))
        if (lt < 0) {
          i = n
        } else {
          this.flushText()
          this.inMarkup = true
          this.kind = Kind.Unknown
          this.markup = []
          this.markupLength = 0
          this.quote = 0
          this.tail = ""
          i = lt + 1
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
          this.fail(this.malformed("text outside the root element", this.consumed))
          return
        }
      }
      return
    }
    const value = this.decode(raw)
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
              this.consumed + i
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
    const terminator = TERMINATOR[this.kind]!
    const window = this.tail + chunk.slice(i)
    // a comment's "-->" may not reuse the dashes of its opening "!--"
    const from = this.kind === Kind.Comment ? Math.max(0, 3 - this.markupLength) : 0
    const at = window.indexOf(terminator, from)
    this.work += window.length
    if (at < 0) {
      this.addMarkup(chunk.slice(i))
      const all = this.tail + chunk.slice(i)
      this.tail = all.slice(Math.max(0, all.length - (terminator.length - 1)))
      return n
    }
    // `at` counts from the start of the tail; the terminator ends in this chunk
    const endInChunk = i + (at - this.tail.length) + terminator.length
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
          this.fail(this.malformed("CDATA outside the root element", this.consumed))
          return
        }
        // markup is "![CDATA[" + content + "]]>"
        const content = markup.slice(8, markup.length - 3)
        if (content.length > this.limits.maxStringLength) {
          this.fail(new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: content.length }))
          return
        }
        if (content.length > 0) this.emit(this.handler.text(content))
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
    if (!NAME.test(qname)) {
      this.fail(this.malformed(`bad element name ${JSON.stringify(qname.slice(0, 40))}`, this.consumed))
      return
    }
    if (this.rootClosed) {
      this.fail(this.malformed("more than one root element", this.consumed))
      return
    }
    const attributes: Array<readonly [string, string]> = []
    const seen = new Set<string>()
    // progress: each iteration consumes one attribute (p moves past its
    // closing quote) or ends the loop
    while (p < n) {
      while (p < n && isSpace(body.charCodeAt(p))) p++
      if (p >= n) break
      const eq = body.indexOf("=", p)
      if (eq < 0) {
        this.fail(this.malformed(`attribute without a value in <${qname}>`, this.consumed))
        return
      }
      const name = body.slice(p, eq).trim()
      let q = eq + 1
      while (q < n && isSpace(body.charCodeAt(q))) q++
      const quote = body.charCodeAt(q)
      if (quote !== 0x22 && quote !== 0x27) {
        this.fail(this.malformed(`unquoted attribute value in <${qname}>`, this.consumed))
        return
      }
      const close = body.indexOf(String.fromCharCode(quote), q + 1)
      if (close < 0 || !NAME.test(name)) {
        this.fail(this.malformed(`bad attribute in <${qname}>`, this.consumed))
        return
      }
      if (seen.has(name)) {
        this.fail(this.malformed(`duplicate attribute ${name} in <${qname}>`, this.consumed))
        return
      }
      seen.add(name)
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
      const value = this.decode(raw)
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
      this.fail(this.malformed(`end tag </${qname}> without a start tag`, this.consumed))
      return
    }
    if (top !== qname) {
      this.fail(this.malformed(`end tag </${qname}> does not match <${top}>`, this.consumed))
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
  private decode(raw: string): string | undefined {
    let amp = raw.indexOf("&")
    if (amp < 0) return raw
    const parts: Array<string> = []
    let from = 0
    // progress: amp moves past each reference (to the next '&' after its ';')
    while (amp >= 0) {
      const semi = raw.indexOf(";", amp + 1)
      if (semi < 0 || semi - amp > 12) {
        this.fail(this.malformed("bad entity reference", this.consumed))
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
        const digits = name.slice(hex ? 2 : 1)
        const valid = hex ? /^[0-9A-Fa-f]{1,6}$/.test(digits) : /^[0-9]{1,7}$/.test(digits)
        const code = valid ? Number.parseInt(digits, hex ? 16 : 10) : -1
        if (
          code === 0x9 || code === 0xA || code === 0xD || (code >= 0x20 && code <= 0xD7FF) ||
          (code >= 0xE000 && code <= 0xFFFD) || (code >= 0x10000 && code <= 0x10FFFF)
        ) ch = String.fromCodePoint(code)
      }
      if (ch === undefined) {
        this.fail(this.malformed(`unknown or invalid entity &${name.slice(0, 12)};`, this.consumed))
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

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
 * Limits: element depth (`maxXmlDepth`), attribute values in one tag
 * (`maxXmlAttributes`), and, each on its own, by `maxStringLength`: a text
 * run (and collected text), an attribute value, the markup before, between
 * and after a tag's values (its names, `=`, whitespace and a `/`), and the
 * content of a comment, processing instruction or CDATA section. A tag is
 * checked as it is scanned, part by part, so a tag split across pieces
 * holds at most these.
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
 * A handler can ask, when an element opens, for its subtree to be skipped
 * (no events from inside it, nor its end tag), for its text to be collected
 * (handed to `close`, bounded as a whole by `maxStringLength`), or for its
 * text to be left out of collected text. Skipped content is still read and
 * checked like any other.
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
 * The element being opened, as `XmlHandler.open` sees it: the handler may
 * call one of these before it returns, to choose what the tokenizer does
 * with the element's content.
 */
export interface XmlElement {
  /**
   * Skips the element's subtree: no events for anything inside it, nor for
   * its end tag. The content is still read and checked (well-formedness,
   * limits), so the same input fails the same way. Text inside it is not
   * part of an enclosing element's collected text.
   */
  readonly skip: () => void
  /**
   * Collects the element's text: every text run and CDATA section inside
   * it, nested elements included (except skipped subtrees), joined and
   * handed to `close`. Their total, plus `prior` (the length of text
   * already collected for the same value, e.g. earlier runs of a rich
   * string), may not exceed `maxStringLength`. Inside an element whose text
   * is collected already, this does nothing: the text goes to that element.
   */
  readonly collect: (prior?: number) => void
  /**
   * Leaves the element's text, its descendants' included, out of collected
   * text: an enclosing element's, and any asked for inside it (their end
   * tags get `undefined`). Unlike `skip`, the events still come.
   */
  readonly omitText: () => void
}

/**
 * Receives tokens. Names are local names (any `prefix:` removed); attribute
 * names are kept as written (e.g. `r:id`). `close` gets the element's
 * collected text if `open` asked for it (`XmlElement.collect`), else
 * `undefined`. Consecutive `text` calls belong to the same text run (text
 * and CDATA arrive separately); text in a skipped subtree doesn't come. A
 * returned error stops tokenizing.
 */
export interface XmlHandler {
  readonly open: (name: string, attributes: XmlAttributes, element: XmlElement) => SpreadsheetError | undefined | void
  readonly close: (name: string, text: string | undefined) => SpreadsheetError | undefined | void
  readonly text?: (value: string) => SpreadsheetError | undefined | void
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
 * Markup (anything that begins with `<`) that continues past the end of a
 * piece. `parts` collects everything after the `<` (`length` characters in
 * all), across pieces.
 * - `Undecided`: the first characters don't tell the construct yet (`<!-`
 *   could become a comment or a DTD).
 * - `Tag`: a start or end tag; its state is the tokenizer's `TagScan`.
 * - `Delimited`: a processing instruction, comment or CDATA section; `tail`
 *   holds its last characters past the opener, for a terminator split
 *   across pieces.
 */
type Markup = Undecided | Tag | Delimited
type Undecided = { readonly _tag: "Undecided"; readonly parts: Array<string>; length: number }
type Tag = { readonly _tag: "Tag" }
type Delimited = {
  readonly _tag: "Delimited"
  readonly kind: Delimiter
  readonly parts: Array<string>
  length: number
  tail: string
}

/** Where the tokenizer is between pieces: in text, or inside markup. */
type Mode = { readonly _tag: "Text" } | Markup

/** Shared, so that text and tags cost no allocation. */
const TEXT: Mode = { _tag: "Text" }
const TAG: Tag = { _tag: "Tag" }

/**
 * The scan of a start or end tag, carried across pieces (one per
 * tokenizer, reset at each tag).
 * - `start`: a start tag, where quotes matter (in an end tag they don't).
 * - `quote`: the quote character of the attribute value the scan is inside
 *   of, or 0.
 * - `part`: the characters of the current part in earlier pieces. The
 *   parts of a tag are its attribute values (inside the quotes) and the
 *   markup before, between and after them (every other character between
 *   `<` and `>` but the quotes).
 * - `values`: the attribute values begun so far.
 * - `parts`: the tag's characters from earlier pieces; empty while the tag
 *   lies within the current piece, which is then read in place, and
 *   emptied when a joined tag is handled.
 */
interface TagScan {
  start: boolean
  quote: number
  part: number
  values: number
  readonly parts: Array<string>
}

/**
 * What a step of the scan returns: undefined to go on, or the error that
 * stops tokenizing. A handler's result is one already.
 */
type Status = SpreadsheetError | undefined | void

/** What a scan over markup returns: the position to go on from, or the error that stops tokenizing. */
type Scanned = number | SpreadsheetError

/** Prolog: before the root element; Root: inside it; Epilog: after it closed. */
type Phase = "Prolog" | "Root" | "Epilog"

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

export class XmlTokenizer implements XmlElement {
  /** Characters examined in the current piece (each input character a bounded number of times). */
  private work = 0

  private readonly meter: Meter
  private readonly handler: XmlHandler
  private readonly limits: XmlLimits
  private readonly part: string | undefined

  private mode: Mode = TEXT
  private readonly tag: TagScan = { start: false, quote: 0, part: 0, values: 0, parts: [] }

  private text: Array<string> = []
  private textLength = 0

  private readonly stack: Array<string> = []
  /** The depth (stack length) of the element whose subtree is skipped, or 0. */
  private skipAt = 0
  /** The depth of the element whose text is collected, or 0. */
  private collectAt = 0
  /** The depth of the element whose text is left out of collected text, or 0. */
  private omitAt = 0
  /** Its text so far, and that text's length plus the `prior` it was asked with (set when a collection starts). */
  private readonly collected: Array<string> = []
  private collectedLength = 0
  /** What `open` asked for the element it was handed ("none" outside `open`). */
  private request: "none" | "skip" | "collect" | "omit" = "none"
  private requestPrior = 0
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

  skip(): void {
    this.request = "skip"
  }

  collect(prior = 0): void {
    this.request = "collect"
    this.requestPrior = prior
  }

  omitText(): void {
    this.request = "omit"
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
          // NaN when the piece ends after '<'
          const first = chunk.charCodeAt(lt + 1)
          if (lt + 1 < n && first !== 0x21 && first !== 0x3F) {
            const next = this.scanTag(this.beginTag(first !== 0x2F), chunk, lt + 1)
            if (typeof next !== "number") {
              error = next
              break
            }
            i = next
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
    return this.deliverText(value)
  }

  /** Text inside the root element: collected and handed to the handler, unless it is in a skipped subtree. */
  private deliverText(value: string): Status {
    if (this.skipAt !== 0) return undefined
    if (this.collectAt !== 0 && this.omitAt === 0) {
      this.collected.push(value)
      this.collectedLength += value.length
      if (this.collectedLength > this.limits.maxStringLength) {
        return new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual: this.collectedLength })
      }
    }
    return this.handler.text === undefined ? undefined : this.handler.text(value)
  }

  /** Collects characters of an undecided construct or of a comment, PI or CDATA section (bounded by their scans). */
  private addMarkup(mode: Undecided | Delimited, s: string): void {
    mode.parts.push(s)
    mode.length += s.length
  }

  private tooLong(actual: number): LimitExceeded {
    return new LimitExceeded({ limit: "maxStringLength", max: this.limits.maxStringLength, actual })
  }

  /** Scans markup from `chunk[i]`; returns the next position (> i), or the error that stops tokenizing. */
  private scanMarkup(mode: Markup, chunk: string, i: number): Scanned {
    switch (mode._tag) {
      case "Undecided":
        return this.scanUndecided(mode, chunk, i)
      case "Tag":
        return this.scanTag(this.tag, chunk, i)
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
    // only "<!" waits for more characters, so a tag decided here starts at i
    if (c === 0x2F) return this.scanTag(this.beginTag(false), chunk, i)
    if (c === 0x3F) return this.scanDelimited(this.decide(mode, "Pi"), chunk, i)
    if (c !== 0x21) return this.scanTag(this.beginTag(true), chunk, i)
    if (head.startsWith("!--")) return this.scanDelimited(this.decide(mode, "Comment"), chunk, i)
    if (head.startsWith("![CDATA[")) return this.scanDelimited(this.decide(mode, "Cdata"), chunk, i)
    if (head.startsWith("!DOCTYPE")) return this.malformed("DTDs (<!DOCTYPE>) are not supported", this.markupStart)
    if (!("!--".startsWith(head) || "![CDATA[".startsWith(head) || "!DOCTYPE".startsWith(head))) {
      return this.malformed("unsupported declaration", this.markupStart)
    }
    // not decidable yet: take the rest of this piece (it's short) and wait
    const take = Math.min(n, i + Math.max(1, head.length + 1 - mode.length))
    this.work += take - i
    this.addMarkup(mode, chunk.slice(i, take))
    return take
  }

  /** The transition out of `Undecided` to a comment, PI or CDATA section: it takes over the characters collected so far. */
  private decide(mode: Undecided, kind: Delimiter): Delimited {
    const next: Delimited = { _tag: "Delimited", kind, parts: mode.parts, length: mode.length, tail: "" }
    this.mode = next
    return next
  }

  /** Starts the scan of a tag whose first character (after '<') is the next one scanned. */
  private beginTag(start: boolean): TagScan {
    const tag = this.tag
    tag.start = start
    tag.quote = 0
    tag.part = 0
    tag.values = 0
    return tag
  }

  /**
   * The one scan of a start or end tag, from `chunk[i]`: finds the '>' that
   * ends it, outside quotes (quotes only matter in start tags), and bounds
   * its parts as it goes (see `TagScan`): each attribute value and each
   * stretch of markup around them by `maxStringLength`, the values by
   * `maxXmlAttributes`. A part's length is checked where it ends, or at the
   * end of the piece. A tag that ends in this piece and began in it is
   * handled in place, on its range of the piece; one that continues is
   * collected, and joined when it ends. Returns the position after its '>',
   * or the piece's end (`n`).
   */
  private scanTag(tag: TagScan, chunk: string, i: number): Scanned {
    const n = chunk.length
    const start = tag.start
    const max = this.limits.maxStringLength
    let quote = tag.quote
    // where the current part began: before i if it began in an earlier piece
    let from = i - tag.part
    let j = i
    // progress: j increases up to n
    for (; j < n; j++) {
      const c = chunk.charCodeAt(j)
      if (quote !== 0) {
        if (c === quote) {
          if (j - from > max) return this.tooLong(j - from)
          quote = 0
          from = j + 1
        }
      } else if (c === 0x3E) break
      else if (start && (c === 0x22 || c === 0x27)) {
        if (j - from > max) return this.tooLong(j - from)
        if (++tag.values > this.limits.maxXmlAttributes) {
          return new LimitExceeded({ limit: "maxXmlAttributes", max: this.limits.maxXmlAttributes, actual: tag.values })
        }
        quote = c
        from = j + 1
      }
    }
    if (j - from > max) return this.tooLong(j - from)
    this.work += j - i + 1
    if (j === n) {
      tag.quote = quote
      tag.part = n - from
      tag.parts.push(chunk.slice(i))
      this.mode = TAG
      return n
    }
    this.mode = TEXT
    let error: Status
    if (tag.parts.length === 0) {
      this.work += j - i
      error = start ? this.startTag(chunk, i, j) : this.endTag(chunk, i, j)
    } else {
      tag.parts.push(chunk.slice(i, j))
      const markup = tag.parts.join("")
      tag.parts.length = 0
      this.work += markup.length
      error = start ? this.startTag(markup, 0, markup.length) : this.endTag(markup, 0, markup.length)
    }
    return error ?? j + 1
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
    const max = this.limits.maxStringLength
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
        this.addMarkup(mode, chunk.slice(i))
        // what is past the opener is content, but for the start of a
        // terminator that may end in the next piece
        const least = mode.length - opener - (terminator.length - 1)
        if (least > max) return this.tooLong(least)
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
    this.addMarkup(mode, chunk.slice(i, endInChunk))
    const content = mode.length - opener - terminator.length
    if (content > max) return this.tooLong(content)
    return this.finishDelimited(mode) ?? endInChunk
  }

  /** The transition back to `Text`: handles the completed comment, PI or CDATA section. */
  private finishDelimited(mode: Delimited): Status {
    this.mode = TEXT
    const markup = mode.parts.length === 1 ? mode.parts[0]! : mode.parts.join("")
    this.work += markup.length
    const kind = mode.kind
    switch (kind) {
      case "Pi":
      case "Comment":
        return undefined
      case "Cdata": {
        if (this.stack.length === 0) return this.malformed("CDATA outside the root element", this.markupStart)
        // markup is "![CDATA[" + content + "]]>" (its length bounded by the scan)
        const content = markup.slice(8, markup.length - 3)
        return content.length > 0 ? this.deliverText(normalizeLineEnds(content)) : undefined
      }
      default:
        return absurd(kind)
    }
  }

  /**
   * A start tag: `src[from, to)` is everything between its '<' and '>'
   * (a range of a piece, or the whole of a tag joined from several). Its
   * values and the markup around them are within bounds (`scanTag`).
   */
  private startTag(src: string, from: number, to: number): Status {
    let n = to
    let selfClosing = false
    // (before `from` is the '<' or nothing, so an empty tag is not self-closing)
    if (src.charCodeAt(n - 1) === 0x2F) {
      selfClosing = true
      n--
    }
    let p = from
    // progress: p increases up to the end of the name
    while (p < n && !isSpace(src.charCodeAt(p))) p++
    const qname = src.slice(from, p)
    if (!isName(qname)) return this.malformed(`bad element name ${JSON.stringify(qname.slice(0, 40))}`, this.markupStart)
    if (this.phase === "Epilog") return this.malformed("more than one root element", this.markupStart)
    const attributes: Array<readonly [string, string]> = []
    // duplicates: a linear check for the usual few attributes, a Set past that
    let seen: Set<string> | undefined
    // progress: each iteration consumes one attribute (p moves past its
    // closing quote) or ends the loop
    while (p < n) {
      while (p < n && isSpace(src.charCodeAt(p))) p++
      if (p >= n) break
      const eq = src.indexOf("=", p)
      if (eq < 0 || eq >= n) return this.malformed(`attribute without a value in <${qname}>`, this.markupStart)
      const name = src.slice(p, eq).trim()
      let q = eq + 1
      while (q < n && isSpace(src.charCodeAt(q))) q++
      // (src[n] is '/', '>' or past the end: never a quote)
      const quote = src.charCodeAt(q)
      if (quote !== 0x22 && quote !== 0x27) return this.malformed(`unquoted attribute value in <${qname}>`, this.markupStart)
      // (a value whose closing quote is past the tag has a name with a quote
      // in it: the quotes were paired differently by the scan, so the name
      // fails)
      const close = src.indexOf(String.fromCharCode(quote), q + 1)
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
      // (the number of attributes and the length of the value are within
      // bounds: each attribute's value is one the scan counted and measured)
      const raw = src.slice(q + 1, close)
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
    if (this.skipAt === 0) {
      const error = this.handler.open(localName(qname), attributes, this)
      const request = this.request
      this.request = "none"
      if (error !== undefined) return error
      if (request === "skip") this.skipAt = this.stack.length
      else if (request === "collect" && this.collectAt === 0 && this.omitAt === 0) {
        this.collectAt = this.stack.length
        this.collectedLength = this.requestPrior
      } else if (request === "omit" && this.omitAt === 0) this.omitAt = this.stack.length
    }
    return selfClosing ? this.closeElement(qname) : undefined
  }

  /** An end tag: `src[from, to)` is everything between its '<' and '>', starting with '/'. */
  private endTag(src: string, from: number, to: number): Status {
    const qname = src.slice(from + 1, to).trim()
    const top = this.stack[this.stack.length - 1]
    if (top === undefined) return this.malformed(`end tag </${qname}> without a start tag`, this.markupStart)
    if (top !== qname) return this.malformed(`end tag </${qname}> does not match <${top}>`, this.markupStart)
    return this.closeElement(qname)
  }

  private closeElement(qname: string): Status {
    const depth = this.stack.length
    this.stack.pop()
    if (this.stack.length === 0) this.phase = "Epilog"
    if (this.skipAt !== 0) {
      if (depth === this.skipAt) this.skipAt = 0
      return undefined
    }
    if (depth === this.omitAt) this.omitAt = 0
    let text: string | undefined
    if (depth === this.collectAt) {
      const parts = this.collected
      text = parts.length === 0 ? "" : parts.length === 1 ? parts[0]! : parts.join("")
      parts.length = 0
      this.collectAt = 0
    }
    return this.handler.close(localName(qname), text)
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

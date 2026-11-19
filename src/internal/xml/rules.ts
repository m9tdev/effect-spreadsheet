/**
 * Path rules: a part reader names the elements it reads by where they sit
 * in the document, and gets the tokenizer handler that calls it for them.
 * The reader keeps no flags or depth counters to know where it is.
 *
 * A path is a list of steps, each a local name (any `prefix:` removed, as
 * the tokenizer gives names) after `/` or `//`:
 * - `/name`: a child of the element the step before matched; as the first
 *   step, the root element.
 * - `//name`: a descendant at any depth of the element the step before
 *   matched; as the first step, any element, the root included.
 *
 * So `//sheets//sheet` is a `<sheet>` anywhere inside a `<sheets>` that is
 * anywhere, and `/sst/si` an `<si>` child of an `<sst>` root.
 *
 * An element may match several paths. Its `forbid` paths are checked first:
 * the first one it matches stops tokenizing with MalformedFile. Then the
 * `open` of each `on` path it matches runs, in the order the paths are
 * given, and at its end tag each `close`, in the same order. A rule's
 * `open` gets the tokenizer's element, to skip the subtree or to collect
 * or leave out its text; a skipped element's subtree gets no rules, and
 * its end tag none either.
 *
 * Matching keeps one small state per open element: the set of path steps
 * its children may match, as a bit mask (at most 30 steps in all).
 *
 * @internal
 */
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { XmlAttributes, XmlElement, XmlHandler } from "./tokenizer.ts"

/** What a rule returns: nothing to go on, or the error that stops tokenizing. */
export type Handled = SpreadsheetError | undefined | void

/** What to do with the elements one path matches. */
export interface XmlRule {
  /** At the start tag. */
  readonly open?: (attributes: XmlAttributes, element: XmlElement) => Handled
  /** At the end tag: the element's text if `open` asked to collect it (`XmlElement.collect`), else `undefined`. */
  readonly close?: (text: string | undefined) => Handled
}

export interface XmlRules {
  /** The elements read, by path. */
  readonly on: Readonly<Record<string, XmlRule>>
  /** Elements that make the part MalformedFile, by path: the reason given for each. */
  readonly forbid?: Readonly<Record<string, string>>
  /** The error for a forbidden element. */
  readonly malformed: (reason: string) => SpreadsheetError
}

/** One step of a path: a bit in the state masks. */
interface Step {
  readonly name: string
  /** `//`: the step can match at any depth below the step before (or below the document), so it stays in the state. */
  readonly anyDepth: boolean
  /** The step after this one, or -1 for the last step of its path. */
  readonly next: number
  /** For the last step: its rule's index (forbid paths first, then `on` paths); else -1. */
  readonly rule: number
}

/** The most steps the rules may have in all: one bit each in a 32-bit mask, with room to spare. */
export const MAX_STEPS = 30

/**
 * Splits a path into steps (`//a/b` is `[["a", true], ["b", false]]`).
 * A bad path is a programming error, so it throws.
 */
export const parsePath = (path: string): ReadonlyArray<readonly [name: string, anyDepth: boolean]> => {
  const steps: Array<readonly [string, boolean]> = []
  let k = 0
  // progress: each iteration moves k past one step's slashes and name
  while (k < path.length) {
    if (path.charCodeAt(k) !== 0x2F) throw new Error(`bad XML path ${JSON.stringify(path)}: a step starts with / or //`)
    const anyDepth = path.charCodeAt(k + 1) === 0x2F
    const from = k + (anyDepth ? 2 : 1)
    const slash = path.indexOf("/", from)
    const end = slash < 0 ? path.length : slash
    const name = path.slice(from, end)
    if (name.length === 0 || name.includes(":")) throw new Error(`bad XML path ${JSON.stringify(path)}: a step is a local name`)
    steps.push([name, anyDepth])
    k = end
  }
  if (steps.length === 0) throw new Error("bad XML path: empty")
  return steps
}

/** The tokenizer's element, as the rules hand it on: it remembers whether the subtree is skipped. */
class RuleElement implements XmlElement {
  element: XmlElement | undefined
  skipped = false
  skip(): void {
    this.skipped = true
    this.element?.skip()
  }
  collect(prior?: number): void {
    this.element?.collect(prior)
  }
  omitText(): void {
    this.element?.omitText()
  }
}

/** The tokenizer handler that applies `rules`. Make one per part read: it holds the open elements' states. */
export const rulesHandler = (rules: XmlRules): XmlHandler => {
  const forbid = Object.entries(rules.forbid ?? {})
  const on = Object.values(rules.on)
  const paths = [...forbid.map(([path]) => path), ...Object.keys(rules.on)]
  const steps: Array<Step> = []
  let start = 0
  // progress: one path per iteration, one step per inner iteration
  for (let rule = 0; rule < paths.length; rule++) {
    const parsed = parsePath(paths[rule]!)
    if (steps.length + parsed.length > MAX_STEPS) throw new Error(`XML rules: more than ${MAX_STEPS} path steps`)
    start |= 1 << steps.length
    for (let k = 0; k < parsed.length; k++) {
      const last = k === parsed.length - 1
      steps.push({ name: parsed[k]![0], anyDepth: parsed[k]![1], next: last ? -1 : steps.length + 1, rule: last ? rule : -1 })
    }
  }
  const forbidden = (1 << forbid.length) - 1
  // states[d]: the steps the children of the open element at depth d can
  // match (depth 0: the document, whose child is the root element);
  // matched[d]: the rules the element at depth d matched (bit r: rule r)
  const states: Array<number> = [start]
  const matched: Array<number> = [0]
  let depth = 0
  const element = new RuleElement()
  return {
    open: (name, attributes, tokenizerElement) => {
      let state = 0
      let rulesHit = 0
      // progress: each iteration clears the lowest bit left
      for (let rest = states[depth]!; rest !== 0; rest &= rest - 1) {
        const bit = rest & -rest
        const step = steps[31 - Math.clz32(bit)]!
        if (step.anyDepth) state |= bit
        if (step.name === name) {
          if (step.next >= 0) state |= 1 << step.next
          else rulesHit |= 1 << step.rule
        }
      }
      if ((rulesHit & forbidden) !== 0) {
        return rules.malformed(forbid[31 - Math.clz32(rulesHit & -rulesHit)]![1])
      }
      if (rulesHit !== 0) {
        element.element = tokenizerElement
        element.skipped = false
        let error: Handled
        // progress: each iteration clears the lowest bit left (rules in their order)
        for (let rest = rulesHit; rest !== 0 && error === undefined; rest &= rest - 1) {
          const open = on[31 - Math.clz32(rest & -rest) - forbid.length]!.open
          if (open !== undefined) error = open(attributes, element)
        }
        element.element = undefined
        if (error !== undefined) return error
      }
      // a skipped element's end tag doesn't come
      if (rulesHit !== 0 && element.skipped) return undefined
      depth++
      states[depth] = state
      matched[depth] = rulesHit
      return undefined
    },
    close: (_, text) => {
      const rulesHit = matched[depth]!
      depth--
      // progress: each iteration clears the lowest bit left (rules in their order)
      for (let rest = rulesHit; rest !== 0; rest &= rest - 1) {
        const close = on[31 - Math.clz32(rest & -rest) - forbid.length]!.close
        const error = close === undefined ? undefined : close(text)
        if (error !== undefined) return error
      }
      return undefined
    }
  }
}

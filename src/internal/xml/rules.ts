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
 * the first one it matches stops tokenizing with MalformedFile. Then, if it
 * matches a `skip` path, its subtree is skipped: no rules for it, for
 * anything inside it, nor for its end tag (the tokenizer still reads and
 * checks it). Else the `open` of each `on` path it matches runs, in the
 * order the paths are given, and at its end tag each `close`, in the same
 * order.
 *
 * Matching keeps one small state per open element: the set of path steps
 * its children may match, as a bit mask (at most 30 distinct steps, and
 * 30 paths, in all). Paths that start alike share their first steps. The
 * last transition at each depth is kept, so a run of siblings with the
 * same name (the items of a list) costs one comparison each.
 *
 * @internal
 */
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { XmlAttributes, XmlElement, XmlHandler } from "./tokenizer.ts"

/** What a rule returns: nothing to go on, or the error that stops tokenizing. */
export type Handled = SpreadsheetError | undefined | void

/** What a rule may ask of the tokenizer for an element's text (skipping is a `skip` path). */
export type TextRequest = Pick<XmlElement, "collect" | "omitText">

/** What to do with the elements one path matches. */
export interface XmlRule {
  /** At the start tag. */
  readonly open?: (attributes: XmlAttributes, element: TextRequest) => Handled
  /** At the end tag: the element's text if `open` asked to collect it (`XmlElement.collect`), else `undefined`. */
  readonly close?: (text: string | undefined) => Handled
}

export interface XmlRules {
  /** The elements read, by path. */
  readonly on: Readonly<Record<string, XmlRule>>
  /** Elements that make the part MalformedFile, by path: the reason given for each. */
  readonly forbid?: Readonly<Record<string, string>>
  /** Elements whose subtrees are skipped, by path. */
  readonly skip?: ReadonlyArray<string>
  /** The error for a forbidden element. */
  readonly malformed: (reason: string) => SpreadsheetError
}

/**
 * One step of the paths, a bit in the state masks. Paths that start alike
 * share their first steps (`//si` and `//si//t` share `//si`), so an
 * element costs one test per distinct step its parent's state holds.
 */
interface Step {
  readonly name: string
  /** `//`: the step can match at any depth below the step before (or below the document), so it stays in the state. */
  readonly anyDepth: boolean
  /** The steps right after this one (their bits), entered when it matches. */
  next: number
  /** The paths that end with this step (bit r: path r; forbid paths first, then skip paths, then `on` paths). */
  ends: number
}

/** The most distinct steps, and the most paths, the rules may have: one bit each in a 32-bit mask, with room to spare. */
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

/** The tokenizer handler that applies `rules`. Make one per part read: it holds the open elements' states. */
export const rulesHandler = (rules: XmlRules): XmlHandler => {
  const forbid = Object.entries(rules.forbid ?? {})
  const skip = rules.skip ?? []
  const opens = Object.values(rules.on).map((rule) => rule.open)
  const closes = Object.values(rules.on).map((rule) => rule.close)
  const paths = [...forbid.map(([path]) => path), ...skip, ...Object.keys(rules.on)]
  if (paths.length > MAX_STEPS) throw new Error(`XML rules: more than ${MAX_STEPS} paths`)
  const steps: Array<Step> = []
  /** The steps by their place: the step before (-1: the document), then `//` or `/`, then the name. */
  const byPlace = new Map<string, number>()
  let start = 0
  // progress: one path per iteration, one step per inner iteration
  for (let path = 0; path < paths.length; path++) {
    let before = -1
    for (const [name, anyDepth] of parsePath(paths[path]!)) {
      const place = `${before} ${anyDepth ? "//" : "/"}${name}`
      let index = byPlace.get(place)
      if (index === undefined) {
        if (steps.length >= MAX_STEPS) throw new Error(`XML rules: more than ${MAX_STEPS} path steps`)
        index = steps.length
        byPlace.set(place, index)
        steps.push({ name, anyDepth, next: 0, ends: 0 })
        if (before < 0) start |= 1 << index
        else steps[before]!.next |= 1 << index
      }
      before = index
    }
    steps[before]!.ends |= 1 << path
  }
  const forbidden = (1 << forbid.length) - 1
  const skipped = ((1 << skip.length) - 1) << forbid.length
  /** The bit of the first `on` path: `on` path k is bit `first + k`. */
  const first = forbid.length + skip.length
  // per depth d (0: the document, whose child is the root element):
  // states[d], the steps the children of the open element at depth d can
  // match; matched[d], the `on` paths that element matched (bit k: the k-th);
  // lastFrom[d], lastName[d], lastState[d], lastMatched[d]: the last
  // transition from depth d - 1, from that state with that name
  const states: Array<number> = [start]
  const matched: Array<number> = [0]
  const lastFrom: Array<number> = [-1]
  const lastName: Array<string> = [""]
  const lastState: Array<number> = [0]
  const lastMatched: Array<number> = [0]
  let depth = 0
  return {
    open: (name, attributes, element) => {
      const from = states[depth]!
      const at = depth + 1
      let state: number
      let hit: number
      if (lastFrom[at] === from && lastName[at] === name) {
        state = lastState[at]!
        hit = lastMatched[at]!
      } else {
        state = 0
        hit = 0
        // progress: each iteration clears the lowest bit left
        for (let rest = from; rest !== 0; rest &= rest - 1) {
          const bit = rest & -rest
          const step = steps[31 - Math.clz32(bit)]!
          if (step.anyDepth) state |= bit
          if (step.name === name) {
            state |= step.next
            hit |= step.ends
          }
        }
        lastFrom[at] = from
        lastName[at] = name
        lastState[at] = state
        lastMatched[at] = hit
      }
      if (hit !== 0) {
        if ((hit & forbidden) !== 0) return rules.malformed(forbid[31 - Math.clz32(hit & -hit)]![1])
        if ((hit & skipped) !== 0) {
          // its end tag doesn't come: the depth stays
          element.skip()
          return undefined
        }
        // progress: each iteration clears the lowest bit left (paths in their order)
        for (let rest = hit >>> first; rest !== 0; rest &= rest - 1) {
          const open = opens[31 - Math.clz32(rest & -rest)]
          const error = open === undefined ? undefined : open(attributes, element)
          if (error !== undefined) return error
        }
      }
      depth = at
      states[at] = state
      matched[at] = hit >>> first
      return undefined
    },
    close: (_, text) => {
      const hit = matched[depth]!
      depth--
      // progress: each iteration clears the lowest bit left (paths in their order)
      for (let rest = hit; rest !== 0; rest &= rest - 1) {
        const close = closes[31 - Math.clz32(rest & -rest)]
        const error = close === undefined ? undefined : close(text)
        if (error !== undefined) return error
      }
      return undefined
    }
  }
}

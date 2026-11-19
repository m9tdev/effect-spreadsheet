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
 * order. An `open` may also skip its element's subtree (it returns
 * "skip"), when that depends on what was read before: then the element's
 * later opens don't run, nor any close.
 *
 * An element on an `opaque` path is read only by the paths that go through
 * it: the `//` steps from above it don't reach inside it, and the elements
 * inside it that match no path, and inside which none can, are skipped
 * whole, as is everything below a matched element under which no path
 * goes on. So with `//c` opaque and a path `//c/v`: a `<c>` inside the
 * `<c>` isn't matched, its `<v>` children are, and an `<x>` inside it or
 * inside a `<v>` is skipped with its subtree, its text included.
 *
 * Matching keeps one small state per open element: the set of path steps
 * its children may match, as a bit mask (at most 30 distinct steps, and
 * 30 paths, in all), and whether they are inside an opaque element. Paths
 * that start alike share their first steps. The last transition at each
 * depth is kept, so a run of siblings with the same name (the items of a
 * list) costs one comparison each.
 *
 * @internal
 */
import type { SpreadsheetError } from "../../SpreadsheetError.ts"
import type { XmlAttributes, XmlElement, XmlHandler } from "./tokenizer.ts"

/** What a rule returns: nothing to go on, or the error that stops tokenizing. */
export type Handled = SpreadsheetError | undefined | void

/** What a rule's `open` returns: as `Handled`, or "skip" to skip the element's subtree. */
export type Opened = Handled | "skip"

/** What a rule may ask of the tokenizer for an element's text (skipping is a `skip` path, or `open` returning "skip"). */
export type TextRequest = Pick<XmlElement, "collect" | "omitText">

/** What to do with the elements one path matches. */
export interface XmlRule {
  /** At the start tag; "skip" skips the element's subtree: its later opens don't run, nor any close. */
  readonly open?: (attributes: XmlAttributes, element: TextRequest) => Opened
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
  /** Elements read only by the paths through them, by path: what else is inside them is skipped. */
  readonly opaque?: ReadonlyArray<string>
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
  /** The paths that end with this step (bit r: path r; forbid paths first, then skip, opaque and `on` paths). */
  ends: number
}

/** The most distinct steps, and the most paths, the rules may have: one bit each in a 32-bit mask, with room to spare. */
export const MAX_STEPS = 30

/** The bit, past the steps' bits, of a state inside an opaque element. */
const OPAQUE = 1 << MAX_STEPS

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
  const opaque = rules.opaque ?? []
  const opens = Object.values(rules.on).map((rule) => rule.open)
  const closes = Object.values(rules.on).map((rule) => rule.close)
  const paths = [...forbid.map(([path]) => path), ...skip, ...opaque, ...Object.keys(rules.on)]
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
  // the steps as columns: their names, their next steps and ended paths, and the `//` ones as a mask
  const names = steps.map((step) => step.name)
  // a cheap first test before the names are compared: length and first character
  const keys = steps.map((step) => step.name.length * 0x10000 + step.name.charCodeAt(0))
  const nexts = steps.map((step) => step.next)
  const ends = steps.map((step) => step.ends)
  let anyDepth = 0
  // progress: k increases up to steps.length
  for (let k = 0; k < steps.length; k++) if (steps[k]!.anyDepth) anyDepth |= 1 << k
  const forbidden = (1 << forbid.length) - 1
  const skipped = ((1 << skip.length) - 1) << forbid.length
  const opaqued = ((1 << opaque.length) - 1) << (forbid.length + skip.length)
  /** The bit of the first `on` path: `on` path k is bit `first + k`. */
  const first = forbid.length + skip.length + opaque.length
  // per depth d (0: the document, whose child is the root element), one
  // entry each, added when the document first goes that deep:
  // - states[d]: the steps the children of the open element at depth d can
  //   match, with the OPAQUE bit if they are inside an opaque element;
  // - matched[d]: the `on` paths that element matched (bit k: the k-th);
  // - lastFrom[d], lastKey[d], lastName[d] -> lastState[d], lastMatched[d]:
  //   the last transition into depth d (from that state, with that name)
  const states: Array<number> = [start]
  const matched: Array<number> = [0]
  const lastFrom: Array<number> = [-1]
  const lastKey: Array<number> = [0]
  const lastName: Array<string> = [""]
  const lastState: Array<number> = [0]
  const lastMatched: Array<number> = [0]
  let depth = 0
  return {
    open: (name, attributes, element) => {
      const from = states[depth]!
      const at = depth + 1
      if (at === states.length) {
        states.push(0)
        matched.push(0)
        lastFrom.push(-1)
        lastKey.push(0)
        lastName.push("")
        lastState.push(0)
        lastMatched.push(0)
      }
      let state: number
      let hit: number
      const key = name.length * 0x10000 + name.charCodeAt(0)
      if (lastFrom[at] === from && lastKey[at] === key && lastName[at] === name) {
        state = lastState[at]!
        hit = lastMatched[at]!
      } else {
        // the steps the element enters, and the paths it matches
        let entered = 0
        hit = 0
        // progress: each iteration clears the lowest bit left
        for (let rest = from & ~OPAQUE; rest !== 0; rest &= rest - 1) {
          const k = 31 - Math.clz32(rest & -rest)
          if (keys[k] === key && names[k] === name) {
            entered |= nexts[k]!
            hit |= ends[k]!
          }
        }
        // the `//` steps from above stay for the element's children, unless it is opaque
        state = (hit & opaqued) !== 0 ? entered | OPAQUE : (from & (anyDepth | OPAQUE)) | entered
        lastFrom[at] = from
        lastKey[at] = key
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
          const opened = open === undefined ? undefined : open(attributes, element)
          if (opened === "skip") {
            // as for a skip path: no more opens, and its end tag doesn't come
            element.skip()
            return undefined
          }
          if (opened !== undefined) return opened
        }
      } else if (state === OPAQUE) {
        // inside an opaque element, and nothing can match inside it either
        element.skip()
        return undefined
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

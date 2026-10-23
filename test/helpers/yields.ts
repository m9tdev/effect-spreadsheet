/**
 * Counts how often the event loop got control while `run` was in flight: a
 * setImmediate ticker that re-arms itself runs once per turn of the loop.
 * Unlike a wall-clock gap, the count doesn't depend on machine load: code
 * that yields N times lets the ticker run about N times. A count alone can't
 * see one long stretch among many short ones, though (it's only "about once"
 * when the whole run is a single stretch). So callers that can measure their
 * progress pass `progress` (a monotonic counter, e.g. records emitted), and
 * get `maxPerTurn`: the most progress made between two turns of the loop,
 * which bounds the longest synchronous stretch in units of work, again
 * independent of load.
 */
export const countYields = async <A>(
  run: () => Promise<A>,
  progress?: () => number
): Promise<{ value: A; yields: number; maxPerTurn: number; ms: number }> => {
  let yields = 0
  let running = true
  let last = progress?.() ?? 0
  let maxPerTurn = 0
  const sample = () => {
    if (progress === undefined) return
    const now = progress()
    maxPerTurn = Math.max(maxPerTurn, now - last)
    last = now
  }
  const tick = () => {
    if (!running) return
    yields++
    sample()
    setImmediate(tick)
  }
  setImmediate(tick)
  const start = performance.now()
  try {
    const value = await run()
    sample()
    return { value, yields, maxPerTurn, ms: performance.now() - start }
  } finally {
    running = false
  }
}

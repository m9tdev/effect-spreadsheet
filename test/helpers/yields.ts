/**
 * Counts how often the event loop got control while `run` was in flight: a
 * setImmediate ticker that re-arms itself runs once per turn of the loop.
 * Unlike a wall-clock gap, the count doesn't depend on machine load: code
 * that yields N times lets the ticker run about N times, and code that
 * blocks in one synchronous stretch lets it run about once.
 */
export const countYields = async <A>(run: () => Promise<A>): Promise<{ value: A; yields: number; ms: number }> => {
  let yields = 0
  let running = true
  const tick = () => {
    if (!running) return
    yields++
    setImmediate(tick)
  }
  setImmediate(tick)
  const start = performance.now()
  try {
    const value = await run()
    return { value, yields, ms: performance.now() - start }
  } finally {
    running = false
  }
}

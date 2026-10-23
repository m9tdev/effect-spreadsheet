import { spawn } from "node:child_process"
import { describe, expect, it } from "vitest"
import { periods } from "./setup/watchdog-periods.ts"

const WATCHDOG = new URL("./setup/watchdog.ts", import.meta.url).pathname

/** Runs `body` in a fresh node process with the watchdog loaded and a 2 s hang limit (heartbeat 500 ms). */
const run = (body: string, hangMs = 2000) =>
  new Promise<{ code: number | null; signal: NodeJS.Signals | null; ms: number }>((resolve) => {
    const start = performance.now()
    const child = spawn(process.execPath, ["--input-type=module", "-e", `await import(${JSON.stringify(WATCHDOG)}); ${body}`], {
      env: { ...process.env, TEST_HANG_MS: String(hangMs) },
      stdio: "ignore"
    })
    child.on("exit", (code, signal) => resolve({ code, signal, ms: performance.now() - start }))
  })

describe("test watchdog", () => {
  it("kills a process whose event loop is blocked", async () => {
    const result = await run("while (true) {}")
    expect(result.signal).toBe("SIGKILL")
    // 2 s limit; a generous ceiling for a loaded machine
    expect(result.ms).toBeLessThan(30_000)
  }, 40_000)

  it("leaves a process alone that waits asynchronously (longer than the hang limit)", async () => {
    const result = await run("await new Promise((r) => setTimeout(r, 3000))")
    expect([result.code, result.signal]).toEqual([0, null])
  }, 20_000)

  it("the heartbeat comes at least 4 times per hang limit (pure check, no timing)", () => {
    // A child process with a 300 ms limit used to be killed while healthy,
    // because the heartbeat was a fixed 500 ms. Checked on the formula, so
    // machine load can't make it flaky.
    for (const hangMs of [4, 10, 100, 300, 1000, 2000, 60_000, 3_600_000]) {
      const { beatMs, checkMs } = periods(hangMs)
      expect(beatMs * 4, `beat for ${hangMs}`).toBeLessThanOrEqual(hangMs)
      expect(checkMs * 4, `check for ${hangMs}`).toBeLessThanOrEqual(hangMs)
      expect(beatMs).toBeGreaterThanOrEqual(1)
    }
    expect(periods(300).beatMs).toBe(75)
    expect(periods(60_000)).toEqual({ beatMs: 500, checkMs: 250 })
  })
})

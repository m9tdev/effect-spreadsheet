import { spawn } from "node:child_process"
import { describe, expect, it } from "vitest"
import { periods } from "./setup/watchdog-periods.ts"

const WATCHDOG = new URL("./setup/watchdog.ts", import.meta.url).pathname

/**
 * Runs `body` in a fresh node process with the watchdog loaded and a 2 s hang
 * limit (heartbeat 500 ms). If the child is still running after `deadlineMs`
 * (say, the watchdog is broken and the body loops), the test kills it, so it
 * can never outlive the test run; `killedByTest` tells the two kills apart.
 */
const run = (body: string, hangMs = 2000, deadlineMs = 30_000) =>
  new Promise<
    { code: number | null; signal: NodeJS.Signals | null; ms: number; killedByTest: boolean; stdout: string }
  >((resolve) => {
    const start = performance.now()
    const child = spawn(process.execPath, ["--input-type=module", "-e", `await import(${JSON.stringify(WATCHDOG)}); ${body}`], {
      env: { ...process.env, TEST_HANG_MS: String(hangMs) },
      stdio: ["ignore", "pipe", "ignore"]
    })
    let stdout = ""
    child.stdout.on("data", (d) => (stdout += String(d)))
    let killedByTest = false
    const deadline = setTimeout(() => {
      killedByTest = true
      child.kill("SIGKILL")
    }, deadlineMs)
    child.on("exit", (code, signal) => {
      clearTimeout(deadline)
      resolve({ code, signal, ms: performance.now() - start, killedByTest, stdout })
    })
  })

describe("test watchdog", () => {
  it("kills a process whose event loop is blocked", async () => {
    const result = await run("while (true) {}")
    expect(result.signal).toBe("SIGKILL")
    // the watchdog did it (2 s limit), not the helper's 30 s deadline
    expect(result.killedByTest).toBe(false)
  }, 40_000)

  it("a child the watchdog doesn't stop is killed at the test's deadline, never left running", async () => {
    // a limit far beyond the deadline stands in for a broken watchdog
    const result = await run("while (true) {}", 1_000_000_000, 1000)
    expect(result.killedByTest).toBe(true)
    expect(result.signal).toBe("SIGKILL")
  }, 20_000)

  it("leaves a process alone that waits asynchronously (longer than the hang limit)", async () => {
    // the child's deadline (15 s) stays inside this test's timeout (20 s)
    const result = await run("await new Promise((r) => setTimeout(r, 3000))", 2000, 15_000)
    expect([result.code, result.signal, result.killedByTest]).toEqual([0, null, false])
  }, 20_000)

  it("this test process runs the watchdog with periods(HANG_MS)", () => {
    const settings = (globalThis as { [key: symbol]: unknown })[Symbol.for("effect-spreadsheet/test-watchdog-settings")] as
      | { hangMs: number; beatMs: number; checkMs: number }
      | undefined
    const hangMs = Number(process.env["TEST_HANG_MS"] ?? 60_000)
    expect(settings).toEqual({ hangMs, ...periods(hangMs) })
  })

  it("a process with a 300 ms limit runs the watchdog with periods(300)", async () => {
    // through the same helper, so the child has a deadline too
    const result = await run(
      `console.log(JSON.stringify(globalThis[Symbol.for("effect-spreadsheet/test-watchdog-settings")]))`,
      300,
      15_000
    )
    expect([result.code, result.killedByTest]).toEqual([0, false])
    expect(JSON.parse(result.stdout)).toEqual({ hangMs: 300, ...periods(300) })
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

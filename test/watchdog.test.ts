import { spawn } from "node:child_process"
import { describe, expect, it } from "vitest"

const WATCHDOG = new URL("./setup/watchdog.ts", import.meta.url).pathname

/** Runs `body` in a fresh node process with the watchdog loaded and a 300 ms hang limit. */
const run = (body: string) =>
  new Promise<{ code: number | null; signal: NodeJS.Signals | null; ms: number }>((resolve) => {
    const start = performance.now()
    const child = spawn(process.execPath, ["--input-type=module", "-e", `await import(${JSON.stringify(WATCHDOG)}); ${body}`], {
      env: { ...process.env, TEST_HANG_MS: "300" },
      stdio: "ignore"
    })
    child.on("exit", (code, signal) => resolve({ code, signal, ms: performance.now() - start }))
  })

describe("test watchdog", () => {
  it("kills a process whose event loop is blocked", async () => {
    const result = await run("while (true) {}")
    expect(result.signal).toBe("SIGKILL")
    expect(result.ms).toBeLessThan(10_000)
  }, 20_000)

  it("leaves a process alone that waits asynchronously", async () => {
    const result = await run("await new Promise((r) => setTimeout(r, 1000))")
    expect([result.code, result.signal]).toEqual([0, null])
  }, 20_000)
})

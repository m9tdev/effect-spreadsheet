import { spawn } from "node:child_process"
import { describe, expect, it } from "vitest"

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
    expect(result.ms).toBeLessThan(10_000)
  }, 20_000)

  it("leaves a process alone that waits asynchronously (longer than the hang limit)", async () => {
    const result = await run("await new Promise((r) => setTimeout(r, 3000))")
    expect([result.code, result.signal]).toEqual([0, null])
  }, 20_000)

  it("with a short limit (300 ms) the heartbeat keeps up: a healthy process survives", async () => {
    // the heartbeat used to be a fixed 500 ms, longer than this limit
    const result = await run("await new Promise((r) => setTimeout(r, 1500))", 300)
    expect([result.code, result.signal]).toEqual([0, null])
  }, 20_000)
})

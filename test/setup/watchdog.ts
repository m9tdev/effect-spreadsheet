/**
 * Hang watchdog for every test process (vitest setupFiles). A worker thread
 * expects a heartbeat from the main thread's event loop; if none arrives for
 * HANG_MS (a synchronous infinite loop), it kills the test process, so the
 * run fails instead of hanging. Async waits are covered by testTimeout.
 * Times use the monotonic performance.now(), so clock changes do not matter.
 */
import { Worker } from "node:worker_threads"

const HANG_MS = Number(process.env["TEST_HANG_MS"] ?? 60_000)

const key = Symbol.for("effect-spreadsheet/test-watchdog")
const g = globalThis as { [key]?: Worker }

if (g[key] === undefined) {
  const source = `
    const { parentPort, workerData } = process.getBuiltinModule("node:worker_threads")
    let last = performance.now()
    parentPort.on("message", () => { last = performance.now() })
    setInterval(() => {
      if (performance.now() - last > workerData.hangMs) {
        process._rawDebug("[watchdog] the test process' event loop was blocked for " + workerData.hangMs +
          " ms (synchronous hang?): killing it")
        process.kill(process.pid, "SIGKILL")
      }
    }, 250)
  `
  const watchdog = new Worker(source, { eval: true, workerData: { hangMs: HANG_MS } })
  watchdog.unref()
  setInterval(() => watchdog.postMessage(0), 500).unref()
  g[key] = watchdog
}

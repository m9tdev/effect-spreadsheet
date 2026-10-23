import { defineConfig } from "vitest/config"

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    // A hang must fail the run: async waits hit these timeouts, synchronous
    // loops are caught by the watchdog (test/setup/watchdog.ts), fuzz cases
    // run in a worker that is terminated after its budget (test/fuzz).
    testTimeout: 60_000,
    hookTimeout: 60_000,
    pool: "forks",
    setupFiles: ["test/setup/watchdog.ts"]
  }
})

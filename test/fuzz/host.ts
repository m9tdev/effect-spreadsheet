/** Runs fuzz tasks in a worker thread with a time budget (see worker.ts). */
import { Worker } from "node:worker_threads"
import type { FuzzTask } from "./tasks.ts"

export type FuzzResult =
  | { readonly kind: "value" | "typed" }
  | { readonly kind: "defect"; readonly detail: string }
  | { readonly kind: "hang"; readonly detail: string }

const WORKER_URL = new URL("./worker.ts", import.meta.url)

export class FuzzHost {
  private worker: Worker | undefined
  private nextId = 0
  private readonly budgetMs: number

  constructor(budgetMs: number) {
    this.budgetMs = budgetMs
  }

  private get(): Worker {
    if (this.worker === undefined) this.worker = new Worker(WORKER_URL)
    return this.worker
  }

  /** A warm-up task so module loading doesn't count against the first budget. */
  async start(): Promise<void> {
    await this.run({ op: "detect", bytes: new Uint8Array(0) }, 60_000)
  }

  run(task: FuzzTask, budgetMs = this.budgetMs): Promise<FuzzResult> {
    const worker = this.get()
    const id = this.nextId++
    return new Promise((resolve) => {
      const done = (result: FuzzResult) => {
        clearTimeout(timer)
        worker.off("message", onMessage)
        worker.off("error", onError)
        worker.off("exit", onExit)
        resolve(result)
      }
      const timer = setTimeout(() => {
        // A synchronous loop can't be interrupted: kill the thread.
        void worker.terminate()
        this.worker = undefined
        done({ kind: "hang", detail: `no result within ${budgetMs} ms for ${describe(task)}` })
      }, budgetMs)
      const onMessage = (reply: { id: number; kind: FuzzResult["kind"]; detail?: string }) => {
        if (reply.id !== id) return
        done(reply.kind === "value" || reply.kind === "typed"
          ? { kind: reply.kind }
          : { kind: "defect", detail: reply.detail ?? "" })
      }
      const onError = (error: unknown) => {
        this.worker = undefined
        done({ kind: "defect", detail: `worker error: ${String(error)}` })
      }
      const onExit = (code: number) => {
        this.worker = undefined
        done({ kind: "defect", detail: `worker exited with ${code}` })
      }
      worker.on("message", onMessage)
      worker.on("error", onError)
      worker.on("exit", onExit)
      worker.postMessage({ id, task })
    })
  }

  async stop(): Promise<void> {
    const worker = this.worker
    this.worker = undefined
    if (worker !== undefined) await worker.terminate()
  }
}

const describe = (task: FuzzTask): string => {
  const bytes = task.op === "csv" ? task.chunks.reduce((n, c) => n + c.length, 0) : task.bytes.length
  return `${task.op} of ${bytes} bytes`
}

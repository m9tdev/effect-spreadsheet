/**
 * Worker thread that runs fuzz tasks. The host terminates it when a task
 * overruns its time budget, so even a synchronous infinite loop becomes a
 * test failure instead of a hang.
 */
import { Effect, Exit, Option } from "effect"
import { parentPort } from "node:worker_threads"
import { classify } from "../helpers/classify.ts"
import { effectOf, type FuzzTask } from "./tasks.ts"

parentPort!.on("message", async (message: { id: number; task: FuzzTask }) => {
  let reply: { id: number; kind: string; detail?: string; tag?: string; reason?: string }
  try {
    const exit = await Effect.runPromise(Effect.exit(effectOf(message.task)))
    reply = { id: message.id, ...classify(exit) }
    if (reply.kind === "typed") {
      // the error's tag and reason, so tests can assert on typed results
      const error = Option.getOrUndefined(Exit.findErrorOption(exit)) as { _tag?: string; reason?: string } | undefined
      if (error?._tag !== undefined) reply.tag = error._tag
      if (typeof error?.reason === "string") reply.reason = error.reason
    }
  } catch (error) {
    reply = { id: message.id, kind: "defect", detail: `runPromise threw: ${String(error)}` }
  }
  parentPort!.postMessage(reply)
})

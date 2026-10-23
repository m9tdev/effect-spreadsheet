import { Cause, Exit, Option } from "effect"
import { isSpreadsheetError } from "../../src/SpreadsheetError.ts"

/**
 * "value" for success, "typed" for a SpreadsheetError failure, "defect" for
 * anything else (die, interrupt, untyped failure) with the pretty cause.
 */
export const classify = <A, E>(exit: Exit.Exit<A, E>): { kind: "value" | "typed" | "defect"; detail?: string } => {
  if (Exit.isSuccess(exit)) return { kind: "value" }
  const cause = exit.cause
  if (Cause.hasDies(cause) || Cause.hasInterrupts(cause)) return { kind: "defect", detail: Cause.pretty(cause) }
  const error = Option.getOrUndefined(Exit.findErrorOption(exit))
  return isSpreadsheetError(error) ? { kind: "typed" } : { kind: "defect", detail: Cause.pretty(cause) }
}

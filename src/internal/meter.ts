/**
 * The work meter: the readers report the work they do to it, each in its
 * own unit, once per bounded piece of work, never once per unit: hot loops
 * keep a plain count and hand it over at the end of the piece.
 *
 * It's taken from context where the work runs, as the limits are: a read
 * gets it when it runs, and so does each run of a sheet's rows or details.
 * Without one it is `noMeter`, which ignores the counts. Tests provide a
 * counting meter to check that the work grows linearly with the input and
 * stays bounded between event-loop turns.
 *
 * @internal
 */
import { Context } from "effect"

/**
 * What the work is counted in:
 * - `bytes`: bytes inflated (or copied, for stored entries) out of a zip;
 * - `sectors`: compound-file sectors and mini sectors walked;
 * - `batches`: batches of the XLS globals walk;
 * - `steps`: steps of the XLS sheet readers' loops (records, cells, rows,
 *   queue and sort steps);
 * - `records`: BIFF12 records read;
 * - `characters`: characters the XML tokenizer examined;
 * - `probes`: lookups made allocating record keys.
 */
export type WorkUnit = "bytes" | "sectors" | "batches" | "steps" | "records" | "characters" | "probes"

export interface Meter {
  /** Records `count` more units of work. */
  readonly add: (unit: WorkUnit, count: number) => void
}

/** A meter that ignores the counts. */
export const noMeter: Meter = { add: () => {} }

/** The meter in effect; `noMeter` unless one is provided. */
export const WorkMeter: Context.Reference<Meter> = Context.Reference<Meter>(
  "effect-spreadsheet/internal/WorkMeter",
  { defaultValue: () => noMeter }
)

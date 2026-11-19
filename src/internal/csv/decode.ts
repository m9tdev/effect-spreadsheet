/**
 * Bytes -> text for delimited text: the shared incremental decoder, set up
 * for the read's encoding option.
 *
 * @internal
 */
import { TextInput } from "../text.ts"
import type { Decoding } from "./options.ts"

/**
 * A decoder for one run of the rows. A BOM always wins (it is dropped, and
 * its encoding is used even when another one is requested). With `"auto"`
 * the encoding is detected from the first bytes; invalid bytes become
 * U+FFFD, never an error.
 */
export const csvText = (decoding: Decoding): TextInput<never> =>
  new TextInput(
    decoding === "auto"
      ? { detect: "bom-utf16-utf8-1252", fatal: false }
      : { detect: "bom-only", encoding: decoding.encoding, fatal: false }
  )

/**
 * Opening a zip package workbook (XLSX, XLSB): the walk both formats share,
 * from the package relationships to a `Workbook`. A format gives its part
 * readers (`PackageFormat`); this module finds the workbook part, reads the
 * shared strings and styles, resolves each sheet to its part once, at
 * open, and builds the workbook from one source per sheet.
 *
 * @internal
 */
import { Effect, type Stream } from "effect"
import { MalformedFile, type SpreadsheetError } from "../SpreadsheetError.ts"
import type { Limits } from "../SpreadsheetLimits.ts"
import type { DefinedName, Row, SheetDetails, Visibility, Workbook } from "../Workbook.ts"
import { type CellStyles, NO_STYLES } from "./numFmt.ts"
import { lookup, type Named, packageRels, partRels, readers, type Rels } from "./opc.ts"
import { fromSources, type SheetSource } from "./workbook.ts"
import type { Archive, LocatedEntry } from "./zip/archive.ts"

/** A `MalformedFile` maker for one package format. */
export const malformedIn = (format: "xlsx" | "xlsb") => (reason: string, part?: string, offset?: number) =>
  new MalformedFile({
    format,
    reason,
    ...(part === undefined ? {} : { part }),
    ...(offset === undefined ? {} : { offset })
  })

/** A sheet as the workbook part lists it. */
export interface SheetEntry {
  readonly name: string
  readonly visibility: Visibility
  /** The relationship that leads to the sheet's part. */
  readonly relationshipId: string | undefined
}

/**
 * A sheet's visibility from its state: XLSX's `state` attribute
 * (`hidden`, `veryHidden`) or XLSB's code (1 hidden, 2 very hidden).
 * Anything else is visible.
 */
export const sheetVisibility = (state: string | number | undefined): Visibility =>
  state === "hidden" || state === 1 ? "hidden" : state === "veryHidden" || state === 2 ? "veryHidden" : "visible"

/** What a sheet's rows are read with, besides its bytes. */
export interface SheetContext {
  readonly part: string
  readonly strings: ReadonlyArray<string> | undefined
  readonly styles: CellStyles
  readonly date1904: boolean
  readonly limits: Limits
}

/** A package format's own part readers. */
export interface PackageFormat<W extends { readonly sheets: ReadonlyArray<SheetEntry>; readonly date1904: boolean }> {
  readonly format: "xlsx" | "xlsb"
  readonly malformed: (reason: string, part?: string) => MalformedFile
  readonly readWorkbook: (archive: Archive, entry: LocatedEntry, limits: Limits) => Effect.Effect<W, SpreadsheetError>
  readonly readStrings: (archive: Archive, entry: LocatedEntry, limits: Limits) => Effect.Effect<ReadonlyArray<string>, SpreadsheetError>
  readonly readStyles: (archive: Archive, entry: LocatedEntry, limits: Limits) => Effect.Effect<CellStyles, SpreadsheetError>
  /** The defined names, read after the strings and styles; `rels` are the workbook part's. */
  readonly definedNames: (
    workbook: W,
    rels: Rels,
    archive: Archive,
    limits: Limits
  ) => Effect.Effect<ReadonlyArray<DefinedName>, SpreadsheetError>
  readonly rows: (bytes: Stream.Stream<Uint8Array, SpreadsheetError>, context: SheetContext) => Stream.Stream<Row, SpreadsheetError>
  readonly details: (
    bytes: Stream.Stream<Uint8Array, SpreadsheetError>,
    part: string,
    limits: Limits
  ) => Effect.Effect<SheetDetails, SpreadsheetError>
}

/**
 * Where a sheet's cells are, resolved once at open: its worksheet part, a
 * worksheet part the package lacks, no cells at all (a chart, dialog or
 * macro sheet), or why it can't be read.
 */
type SheetPart =
  | { readonly _tag: "Part"; readonly entry: LocatedEntry }
  | { readonly _tag: "Missing"; readonly name: string }
  | Extract<SheetSource, { readonly _tag: "NoCells" | "Broken" }>

const NO_CELLS: SheetSource = { _tag: "NoCells" }

/**
 * Where each format keeps its workbook part when the package relationships
 * name none (or there are none).
 */
export const fixedWorkbook = { xlsx: "xl/workbook.xml", xlsb: "xl/workbook.bin" } as const

/**
 * The workbook part of a package: the target of the package's
 * `officeDocument` relationship, else `fallback`. It may be missing from the
 * archive; reading the workbook then fails.
 */
export const workbookPart = (archive: Archive, fallback: string, limits: Limits): Effect.Effect<Named, SpreadsheetError> =>
  Effect.map(packageRels(archive, limits), (rels) => rels.named("officeDocument") ?? lookup(archive, fallback))

/**
 * Opens the workbook of a package from its workbook part (`workbookPart`):
 * the part, its relationships, the shared strings, styles and defined
 * names, and one source per sheet. A sheet whose part is missing or can't
 * be named still opens; reading it fails.
 */
export const readPackage = <W extends { readonly sheets: ReadonlyArray<SheetEntry>; readonly date1904: boolean }>(
  archive: Archive,
  workbookTarget: Named,
  limits: Limits,
  f: PackageFormat<W>
): Effect.Effect<Workbook, SpreadsheetError> =>
  Effect.gen(function*() {
    if (workbookTarget._tag === "Missing") return yield* f.malformed(`workbook part ${workbookTarget.name} is missing`)
    const workbookEntry = workbookTarget.entry
    const workbook = yield* f.readWorkbook(archive, workbookEntry, limits)
    const rels = yield* partRels(archive, workbookEntry, limits)
    const broken = (reason: string): SheetPart => ({ _tag: "Broken", error: f.malformed(reason, workbookEntry.name) })

    const sheetPart = (sheet: SheetEntry): SheetPart => {
      const rel = sheet.relationshipId === undefined ? undefined : rels.byId(sheet.relationshipId)
      if (rel === undefined) return broken(`sheet ${JSON.stringify(sheet.name)} has no part`)
      const target = rel.target
      if (target._tag === "External") return broken(`sheet ${JSON.stringify(sheet.name)} has an external target`)
      if (target._tag === "Invalid") return broken(`sheet ${JSON.stringify(sheet.name)} has a bad target: ${target.reason}`)
      if (rel.kind !== "worksheet") return NO_CELLS
      return target
    }
    const sheets = workbook.sheets.map((sheet) => ({ sheet, part: sheetPart(sheet) }))

    const stringsEntry = rels.part("sharedStrings")
    const stylesEntry = rels.part("styles")
    const strings = stringsEntry === undefined ? undefined : yield* f.readStrings(archive, stringsEntry, limits)
    const styles: CellStyles = stylesEntry === undefined ? NO_STYLES : yield* f.readStyles(archive, stylesEntry, limits)
    const definedNames = yield* f.definedNames(workbook, rels, archive, limits)

    // Each sheet reads its part as its own reader: reading it again is free,
    // and sheets that share a part (Excel never writes this) each pay for it.
    const sourceOf = (part: SheetPart, index: number): SheetSource => {
      switch (part._tag) {
        case "NoCells":
        case "Broken":
          return part
        case "Missing":
          return { _tag: "Broken", error: f.malformed(`sheet part ${part.name} is missing`, workbookEntry.name) }
        case "Part": {
          const entry = part.entry
          const reader = readers.sheet(index)
          return {
            _tag: "Cells",
            rows: (effective) =>
              f.rows(archive.read(entry, reader, effective), {
                part: entry.name,
                strings,
                styles,
                date1904: workbook.date1904,
                limits: effective
              }),
            details: (effective) => f.details(archive.read(entry, reader, effective), entry.name, effective)
          }
        }
      }
    }

    return fromSources({
      format: f.format,
      sheets: sheets.map(({ sheet, part }, index) => ({ name: sheet.name, visibility: sheet.visibility, source: sourceOf(part, index) })),
      definedNames,
      limits
    })
  })

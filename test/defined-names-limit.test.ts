/** maxDefinedNames: XLS NAME records and XLSX definedName elements past the limit fail with LimitExceeded. */
import { Effect, Exit } from "effect"
import { describe, expect, it } from "vitest"
import type { LimitOverrides } from "../src/SpreadsheetLimits.ts"
import * as SpreadsheetLimits from "../src/SpreadsheetLimits.ts"
import * as Xls from "../src/xls/index.ts"
import * as Xlsx from "../src/xlsx/index.ts"
import { failureOf } from "./helpers/run.ts"
import * as X from "./helpers/xls.ts"
import { xlsx } from "./helpers/xlsx.ts"

const xlsNames = (count: number) =>
  X.xls({
    globals: [X.externSheet([0]), ...Array.from({ length: count }, (_, k) => X.name(`Name${k}`, X.ref3d(0, k, 0)))],
    sheets: [{ name: "Data", records: [] }]
  })
const xlsxNames = (count: number) =>
  xlsx({
    sheets: [{ name: "Data" }],
    definedNames: Array.from({ length: count }, (_, k) => `<definedName name="Name${k}">Data!$A$${k + 1}</definedName>`).join("")
  })

const readXls = (bytes: Uint8Array, limits: LimitOverrides) => Effect.runPromise(Effect.exit(Xls.read(bytes, { limits })))
const readXlsx = (bytes: Uint8Array, limits: LimitOverrides) => Effect.runPromise(Effect.exit(Xlsx.read(bytes, { limits })))

describe("maxDefinedNames", () => {
  it("has a default far above real workbooks' defined names", () => {
    expect(SpreadsheetLimits.defaults.maxDefinedNames).toBe(500_000)
  })

  it("xls: names up to the limit are read; one more fails with LimitExceeded", async () => {
    const atLimit = await readXls(xlsNames(3), { maxDefinedNames: 3 })
    expect(Exit.isSuccess(atLimit) && atLimit.value.definedNames.map((d) => d.formula)).toEqual(["Data!$A$1", "Data!$A$2", "Data!$A$3"])
    expect(failureOf(await readXls(xlsNames(4), { maxDefinedNames: 3 }))).toMatchObject({
      _tag: "LimitExceeded",
      limit: "maxDefinedNames",
      max: 3,
      actual: 4
    })
  })

  it("xlsx: names up to the limit are read; one more fails with LimitExceeded", async () => {
    const atLimit = await readXlsx(await xlsxNames(3), { maxDefinedNames: 3 })
    expect(Exit.isSuccess(atLimit) && atLimit.value.definedNames.map((d) => d.name)).toEqual(["Name0", "Name1", "Name2"])
    expect(failureOf(await readXlsx(await xlsxNames(4), { maxDefinedNames: 3 }))).toMatchObject({
      _tag: "LimitExceeded",
      limit: "maxDefinedNames",
      max: 3,
      actual: 4
    })
  })
})

// Tests for the sheetjs-quirks.json checks. Run: node --test scripts/corpus/quirks.test.mjs
// Uses the committed fixtures and the committed expected JSONs (no SheetJS needed).
import assert from "node:assert/strict"
import fs from "node:fs"
import path from "node:path"
import test from "node:test"
import { CORPUS_DIR } from "./lib.mjs"
import { checkQuirk, csvFields, loadQuirks } from "./quirks.mjs"

const QUIRKS = loadQuirks()
const docFor = (q) => JSON.parse(fs.readFileSync(path.join(CORPUS_DIR, "expected", q.file + ".json"), "utf8"))
const fileFor = (q) => path.join(CORPUS_DIR, q.file)
const run = (q) => checkQuirk(q, docFor(q), fileFor(q))
const mutate = (q, f) => { const c = structuredClone(q); f(c); return c }

test("every quirk checks out against the oracle output and the fixture bytes", () => {
  for (const q of QUIRKS) assert.doesNotThrow(() => run(q), q.id)
})

test("every quirk has a checkable sheetjs and correct value", () => {
  for (const q of QUIRKS) {
    assert.ok(["hiddenRows", "ref", "cells", "rows", "dimension"].includes(q.field), q.id)
    assert.ok(typeof q.correctCheck === "string", `${q.id}: correctCheck`)
    if (q.field === "cells" || q.field === "rows") for (const v of [q.sheetjs, q.correct]) assert.ok(Array.isArray(v) && v.every((x) => Array.isArray(x) && x.length === 3), `${q.id}: [r, c, v] triples`)
  }
})

test("a wrong sheetjs value fails, for each field type", () => {
  const byField = (f) => QUIRKS.find((q) => q.field === f)
  assert.throws(() => run(mutate(byField("hiddenRows"), (q) => { q.sheetjs = [2] })), /sheetjs: hiddenRows/)
  assert.throws(() => run(mutate(byField("cells"), (q) => { q.sheetjs[0][2] = "something else" })), /sheetjs: cells/)
  assert.throws(() => run(mutate(byField("dimension"), (q) => { q.sheetjs.ref = "A1:XFD1048576" })), /sheetjs: ref/)
  assert.throws(() => run(mutate(byField("rows"), (q) => { q.sheetjs[0][2] = "2030-01-03T12:30:00.000Z" })), /sheetjs: rows/)
  // "no cell" (v null) is checked too
  const empty = QUIRKS.find((q) => q.id === "csv-empty-quoted-field-no-cell")
  assert.throws(() => run(mutate(empty, (q) => { q.sheetjs = [[5, 0, null]] })), /sheetjs: cells/)
})

test("a wrong correct value fails, for each check method", () => {
  for (const method of new Set(QUIRKS.map((q) => q.correctCheck))) {
    const q = QUIRKS.find((x) => x.correctCheck === method)
    const wrong = mutate(q, (c) => {
      if (Array.isArray(c.correct) && Array.isArray(c.correct[0])) c.correct[0][2] = "wrong"
      else if (Array.isArray(c.correct)) c.correct = [0]
      else c.correct.declaredRows = 65536
    })
    assert.throws(() => run(wrong), /correct:/, method)
  }
})

test("csvFields: quotes, escaped quotes, explicit empty field, line endings", () => {
  const m = csvFields('a,"b ""q""",""\r\n"x\ny",\n3')
  assert.equal(m.get("0,0"), "a")
  assert.equal(m.get("0,1"), 'b "q"')
  assert.equal(m.get("0,2"), "") // "" is an explicit empty cell
  assert.equal(m.get("1,0"), "x\ny")
  assert.equal(m.has("1,1"), false) // unquoted empty is no cell
  assert.equal(m.get("2,0"), "3")
})

test("decodeFixture: BOM stripped when present, nothing lost without one", async () => {
  const { decodeFixture } = await import("./quirks.mjs")
  const be = (s) => { const b = Buffer.from(s, "utf16le"); b.swap16(); return b }
  assert.equal(decodeFixture(Buffer.concat([Buffer.from([0xfe, 0xff]), be("a,b")]), "utf-16be"), "a,b")
  assert.equal(decodeFixture(be("a,b"), "utf-16be"), "a,b") // no BOM: the first character survives
  assert.equal(decodeFixture(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from("a,b", "utf16le")]), "utf-16le"), "a,b")
  assert.equal(decodeFixture(Buffer.from("a,b", "utf16le"), "utf-16le"), "a,b")
  assert.equal(decodeFixture(Buffer.from("\ufeffa,b"), "utf-8"), "a,b")
  assert.equal(decodeFixture(Buffer.from("a,b"), "utf-8"), "a,b")
  assert.equal(decodeFixture(Buffer.from([0x80, 0x61]), "windows-1252"), "€a")
})

test("declared-dimension: the cells span is checked, not only the declared size", () => {
  const q = QUIRKS.find((x) => x.correctCheck === "declared-dimension")
  assert.throws(() => run(mutate(q, (c) => { c.correct.cells = "A1:C3" })), /correct: cells span A1:B2/)
  assert.throws(() => run(mutate(q, (c) => { c.correct.declaredCols = 256 })), /correct: declared/)
})

test("xmlText: rich runs joined, phonetic (<rPh>) runs dropped, entities and _xHHHH_ decoded", async () => {
  const { xmlText } = await import("./quirks.mjs")
  assert.equal(xmlText('<t>phonetic</t><rPh sb="0" eb="1"><t>ignored</t></rPh>'), "phonetic")
  assert.equal(xmlText('<r><t>rich </t></r><r><rPr><b/></rPr><t>bold</t></r>'), "rich bold")
  assert.equal(xmlText("<t>&lt;a&gt; &amp; &#x1F363; &#233; _x000D__x000A_</t>"), "<a> & 🍣 é \r\n")
  assert.equal(xmlText("<t/>"), "")
})

test("unknown field or check method is an error, not a silent pass", () => {
  const q = QUIRKS[0]
  assert.throws(() => run(mutate(q, (c) => { c.field = "bogus" })), /unknown field bogus/)
  assert.throws(() => run(mutate(q, (c) => { c.correctCheck = "bogus" })), /unknown correctCheck bogus/)
})

test("date1904 quirk: the correct value comes from the 1904 epoch, and needs date1904 in the workbook", () => {
  const q = QUIRKS.find((x) => x.id === "sheet-to-json-ignores-date1904")
  assert.ok(q)
  assert.doesNotThrow(() => run(q))
  // the same check on the 1900-system twin fails: it doesn't declare date1904
  assert.throws(() => checkQuirk({ ...q, file: "generated/dates.xlsx" }, JSON.parse(fs.readFileSync(path.join(CORPUS_DIR, "expected/generated/dates.xlsx.json"), "utf8")), path.join(CORPUS_DIR, "generated/dates.xlsx")), /date1904/)
})

test("date1904 check: typed and non-numeric cells are skipped, never a bare RangeError", async () => {
  const { zip } = await import("./lib.mjs")
  const os = await import("node:os")
  const ws = '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c><c r="C1"><v>NaN</v></c><c r="D1"><v>1e300</v></c><c r="E1" t="str"><v>x</v></c></row></sheetData></worksheet>'
  const xlsx = zip([["xl/workbook.xml", '<workbook><workbookPr date1904="1"/></workbook>'], ["xl/worksheets/sheet1.xml", ws]])
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "quirk-1904-")), "t.xlsx")
  fs.writeFileSync(file, xlsx)
  try {
    const doc = { sheets: [{ name: "S", rows: [[null, "1900-01-01T00:00:00.000Z"]] }] }
    const q = { id: "t", file: "x", sheet: "S", field: "rows", sheetjs: [[0, 1, "1900-01-01T00:00:00.000Z"]], correct: [[0, 1, "1904-01-02T00:00:00.000Z"], [0, 0, null], [0, 2, null], [0, 3, null], [0, 4, null]], correctCheck: "xlsx-date1904-serial" }
    assert.doesNotThrow(() => checkQuirk(q, doc, file))
    // a wrong expectation fails with the quirk's own error
    assert.throws(() => checkQuirk({ ...q, correct: [[0, 1, "2000-01-01T00:00:00.000Z"]] }, doc, file), /sheetjs-quirks.json t .*correct: 1904-system dates/)
  } finally { fs.rmSync(path.dirname(file), { recursive: true, force: true }) }
})

test("xlsxSheetPart: maps a sheet name to its part through the workbook rels", async () => {
  const { xlsxSheetPart } = await import("./quirks.mjs")
  const { zip } = await import("./lib.mjs")
  const wb = '<workbook xmlns:r="r"><sheets><sheet name="A &amp; B" sheetId="1" r:id="rId2"/><sheet name="Other" sheetId="2" r:id="rId1"/></sheets></workbook>'
  const rels = '<Relationships><Relationship Id="rId1" Type="ws" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="ws" Target="worksheets/sheet7.xml"/></Relationships>'
  const buf = zip([["xl/workbook.xml", wb], ["xl/_rels/workbook.xml.rels", rels]])
  assert.equal(xlsxSheetPart(buf, "A & B"), "xl/worksheets/sheet7.xml")
  assert.equal(xlsxSheetPart(buf, "Other"), "xl/worksheets/sheet1.xml")
  assert.throws(() => xlsxSheetPart(buf, "Missing"), /no sheet "Missing"/)
  // the committed fixtures resolve to their real parts
  assert.equal(xlsxSheetPart(fs.readFileSync(path.join(CORPUS_DIR, "generated/strings.xlsx")), "Strings"), "xl/worksheets/sheet1.xml")
  // a quirk naming a sheet that doesn't exist fails with the quirk's own error
  const q = QUIRKS.find((x) => x.correctCheck === "xlsx-shared-strings")
  const doc = docFor(q)
  doc.sheets[0].name = "Renamed"
  assert.throws(() => checkQuirk({ ...q, sheet: "Renamed" }, doc, fileFor(q)), /sheetjs-quirks.json .*correct: no sheet "Renamed"/)
})

test("date1904 check reads the quirk's sheet, also when it isn't the first one", async () => {
  const { zip } = await import("./lib.mjs")
  const os = await import("node:os")
  const wb = '<workbook xmlns:r="r"><workbookPr date1904="1"/><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Dates" sheetId="2" r:id="rId2"/></sheets></workbook>'
  const rels = '<Relationships><Relationship Id="rId1" Type="ws" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="ws" Target="worksheets/sheet2.xml"/></Relationships>'
  const sheet = (v) => `<worksheet><sheetData><row r="1"><c r="A1"><v>${v}</v></c></row></sheetData></worksheet>`
  const xlsx = zip([["xl/workbook.xml", wb], ["xl/_rels/workbook.xml.rels", rels], ["xl/worksheets/sheet1.xml", sheet(0)], ["xl/worksheets/sheet2.xml", sheet(1)]])
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "quirk-1904-2sheets-")), file = path.join(dir, "t.xlsx")
  fs.writeFileSync(file, xlsx)
  try {
    const doc = { sheets: [{ name: "First", rows: [] }, { name: "Dates", rows: [["1900-01-01T00:00:00.000Z"]] }] }
    const q = { id: "t", file: "x", sheet: "Dates", field: "rows", sheetjs: [[0, 0, "1900-01-01T00:00:00.000Z"]], correct: [[0, 0, "1904-01-02T00:00:00.000Z"]], correctCheck: "xlsx-date1904-serial" }
    assert.doesNotThrow(() => checkQuirk(q, doc, file)) // serial 1 from sheet2.xml, not serial 0 from sheet1.xml
  } finally { fs.rmSync(dir, { recursive: true, force: true }) }
})

// Hidden-row fixture names, shared by hidden-rows/generate.mjs and hidden-rows/manifest.mjs.
export const NAMES = ["hidden-row.xls", "hidden-row.xml", "hidden-row.slk", "hidden-row.xlsb", "hidden-row.xlsx", "plain.xlsx"]
// Misleading-extension copies: same bytes, wrong name. Format detection must ignore the extension.
export const RENAMED = {
  "renamed/hidden-row-xls.csv": "hidden-row.xls",
  "renamed/hidden-row-xml.xlsx": "hidden-row.xml",
  "renamed/hidden-row-slk.csv": "hidden-row.slk",
  "renamed/hidden-row-xlsb.xls": "hidden-row.xlsb",
  "renamed/hidden-row-xlsx.csv": "hidden-row.xlsx",
  "renamed/plain-xlsx.xls": "plain.xlsx",
}

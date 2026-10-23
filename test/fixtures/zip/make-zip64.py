# Regenerates zip64.zip: a small archive written by Python's zipfile with the
# ZIP64 end-of-central-directory record and locator forced on
# (ZIP_FILECOUNT_LIMIT = 0). The classic EOCD still holds the real values;
# test/zip-directory.test.ts patches them to 0xFFFF / 0xFFFFFFFF in memory.
# Usage: python3 test/fixtures/zip/make-zip64.py
import os, zipfile
zipfile.ZIP_FILECOUNT_LIMIT = 0
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "zip64.zip")
with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_STORED) as z:
    for name, data in [
        ("[Content_Types].xml", b"<Types/>"),
        ("xl/workbook.xml", b"<workbook/>"),
        ("xl/worksheets/sheet1.xml", b"<worksheet/>"),
    ]:
        info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
        z.writestr(info, data)

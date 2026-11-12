"""Create an openpyxl workbook with uncached formulas and formatted blank cells."""

import hashlib
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl import Workbook, load_workbook

if len(sys.argv) != 2:
    raise SystemExit("Usage: python3 scripts/uncached-import-fixture.py NEW_XLSX")
output = Path(sys.argv[1])
report = Path(str(output) + ".json")
if output.exists() or report.exists():
    raise SystemExit("Refusing to overwrite evidence")

book = Workbook()
sheet = book.active
sheet.title = "Uncached benchmark"
for row in range(1, 10001):
    sheet.cell(row, 1, row)
    sheet.cell(row, 2, f"=A{row}*2")
    if row % 5 == 0:
        sheet.cell(row, 3).number_format = "0.00"
book.save(output)
restored = load_workbook(output, data_only=False, read_only=False).active
assert sum(restored.cell(row, 2).value == f"=A{row}*2" for row in range(1, 10001)) == 10000
assert sum(restored.cell(row, 3).number_format == "0.00" for row in range(1, 10001)) == 2000
data = output.read_bytes()
report.write_text(
    json.dumps(
        {
            "generator": "openpyxl",
            "openpyxl": openpyxl.__version__,
            "python": sys.version,
            "rows": 10000,
            "uncachedFormulas": 10000,
            "formattedBlankCells": 2000,
            "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
            "independentReadback": "all formulas and number formats checked with openpyxl",
        },
        indent=2,
    )
    + "\n"
)

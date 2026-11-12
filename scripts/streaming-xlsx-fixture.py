"""Create a stored streaming ZIP around an independently authored openpyxl workbook."""

import hashlib
import io
import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation


class Unseekable:
    def __init__(self):
        self.content = bytearray()

    def write(self, data):
        self.content.extend(data)
        return len(data)

    def flush(self):
        pass


if len(sys.argv) != 2:
    raise SystemExit("Usage: python3 scripts/streaming-xlsx-fixture.py NEW_XLSX")
output = Path(sys.argv[1])
report_path = Path(str(output) + ".json")
if output.exists() or report_path.exists():
    raise SystemExit("Refusing to overwrite evidence")

book = Workbook()
book.properties.created = book.properties.modified = datetime(2024, 1, 1)
sheet = book.active
sheet.title = "Independent stream"
sheet["A1"] = 42
sheet["A2"] = "streamed text"
sheet["B1"] = "=A1*2"
validation = DataValidation(type="whole", operator="greaterThan", formula1="0")
sheet.add_data_validation(validation)
validation.add("A1:A2")
authored = io.BytesIO()
book.save(authored)

target = Unseekable()
with zipfile.ZipFile(io.BytesIO(authored.getvalue())) as source:
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_STORED) as zipped:
        for old in source.infolist():
            info = zipfile.ZipInfo(old.filename, date_time=(2024, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED
            content = source.read(old.filename)
            if old.filename == "docProps/core.xml":
                content, substitutions = re.subn(
                    rb"(<dcterms:modified[^>]*>)[^<]+",
                    rb"\g<1>2024-01-01T00:00:00Z",
                    content,
                )
                assert substitutions == 1
            zipped.writestr(info, content)
data = bytes(target.content)
with zipfile.ZipFile(io.BytesIO(data)) as verify:
    entries = verify.infolist()
    assert len(entries) > 1 and all(item.flag_bits & 8 for item in entries)
    assert all(item.compress_type == zipfile.ZIP_STORED for item in entries)
    assert verify.testzip() is None
    for item in entries:
        local_size = int.from_bytes(data[item.header_offset + 18 : item.header_offset + 22], "little")
        assert local_size == 0
        header_name = int.from_bytes(data[item.header_offset + 26 : item.header_offset + 28], "little")
        header_extra = int.from_bytes(data[item.header_offset + 28 : item.header_offset + 30], "little")
        descriptor = item.header_offset + 30 + header_name + header_extra + item.compress_size
        assert data[descriptor : descriptor + 4] == b"PK\x07\x08"
restored = load_workbook(io.BytesIO(data), data_only=False)
assert restored.sheetnames == ["Independent stream"]
assert restored.active["A1"].value == 42
assert restored.active["A2"].value == "streamed text"
assert restored.active["B1"].value == "=A1*2"
assert len(restored.active.data_validations.dataValidation) == 1

output.write_bytes(data)
report_path.write_text(
    json.dumps(
        {
            "generator": "openpyxl + Python zipfile on a nonseekable writer",
            "openpyxl": openpyxl.__version__,
            "python": sys.version,
            "sha256": hashlib.sha256(data).hexdigest(),
            "bytes": len(data),
            "entries": len(entries),
            "descriptor": "signed, stored, local size zero",
            "independentReadback": "openpyxl formula, values, sheet name and validation count passed",
        },
        ensure_ascii=False,
        indent=2,
    )
    + "\n"
)

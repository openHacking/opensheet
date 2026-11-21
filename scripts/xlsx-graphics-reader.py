"""Inspect a synthetic graphics workbook before and after LibreOffice headless save."""

from hashlib import sha256
import json
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory
from zipfile import ZipFile

import openpyxl
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
BENCH = ROOT / "benchmarks"
SOURCE = BENCH / "xlsx-graphics.openpyxl.xlsx"
SAVED = BENCH / "xlsx-graphics.libreoffice.xlsx"
REPORT = BENCH / "xlsx-graphics-reader-2026-11-22.json"


def inspect(path):
    with path.open("rb") as stream:
        book = load_workbook(stream)
        sheet = book["Graphics data"]
        result = {
            "values": [[sheet.cell(row, col).value for col in (1, 2)] for row in (1, 2)],
            "charts": len(sheet._charts),
            "images": len(sheet._images),
        }
        book.close()
    with ZipFile(path) as archive:
        result["chartParts"] = sorted(name for name in archive.namelist() if name.startswith("xl/charts/") and name.endswith(".xml"))
        result["mediaParts"] = sorted(name for name in archive.namelist() if name.startswith("xl/media/") and not name.endswith("/"))
    return result


def artifact(path):
    return {"file": path.name, "sha256": sha256(path.read_bytes()).hexdigest()}


def main():
    version = subprocess.check_output(["soffice", "--version"], text=True, timeout=30).strip()
    before = inspect(SOURCE)
    with TemporaryDirectory() as directory:
        temporary = Path(directory)
        source = temporary / "input" / SOURCE.name
        source.parent.mkdir()
        shutil.copyfile(SOURCE, source)
        output = temporary / "output"
        output.mkdir()
        command = ["soffice", f"-env:UserInstallation={(temporary / 'profile').as_uri()}", "--headless", "--convert-to", "xlsx", "--outdir", str(output), str(source)]
        subprocess.check_output(command, text=True, stderr=subprocess.STDOUT, timeout=60)
        converted = output / SOURCE.name
        after = inspect(converted)
        shutil.copyfile(converted, SAVED)
    report = {
        "reader": version,
        "inspector": f"openpyxl {openpyxl.__version__}",
        "scope": "One synthetic chart and image; headless LibreOffice save inspected by openpyxl and ZIP paths. No visual rendering, Excel or stable LibreOffice coverage.",
        "artifacts": {"source": artifact(SOURCE), "saved": artifact(SAVED)},
        "before": before,
        "after": after,
    }
    REPORT.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps(report, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

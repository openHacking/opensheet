"""Generate and independently read back a chart and image fixture with openpyxl."""

from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image
from PIL import Image as PILImage


def main():
    book = Workbook()
    book.properties.created = datetime(2026, 11, 22, tzinfo=timezone.utc)
    sheet = book.active
    sheet.title = "Graphics data"
    sheet.append(["Sales", 5])
    sheet.append(["Costs", 3])
    chart = BarChart()
    chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=2))
    sheet.add_chart(chart, "D1")
    png = BytesIO()
    PILImage.new("RGB", (2, 2), "red").save(png, format="PNG")
    png.seek(0)
    sheet.add_image(Image(png), "H1")
    raw = BytesIO()
    book.save(raw)
    output = Path(__file__).resolve().parents[1] / "benchmarks" / "xlsx-graphics.openpyxl.xlsx"
    with ZipFile(raw) as source, ZipFile(output, "w") as target:
        for name in sorted(source.namelist()):
            data = source.read(name)
            if name == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]+", rb"\g<1>2026-11-22T00:00:00Z", data)
            target.writestr(ZipInfo(name, (2026, 11, 22, 0, 0, 0)), data, compress_type=ZIP_DEFLATED)
    with output.open("rb") as stream:
        restored = load_workbook(stream)
        sheet = restored.active
        assert [sheet["B1"].value, sheet["B2"].value] == [5, 3]
        assert len(sheet._charts) == len(sheet._images) == 1
        restored.close()
    print(output)


if __name__ == "__main__":
    main()

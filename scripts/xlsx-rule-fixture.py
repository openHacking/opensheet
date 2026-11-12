"""Generate and independently read back the XLSX rule-loss fixture (openpyxl 3.1.5)."""

from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation


def main():
    book = Workbook()
    book.properties.created = datetime(2026, 11, 10, tzinfo=timezone.utc)
    first = book.active
    first.title = "Rule data"
    first.append([5, "red"])
    second = book.create_sheet("Other rules")
    second.append([6])
    for sheet, target, kind, formula in (
        (first, "A1:A20", "whole", "0"),
        (first, "B1:B20", "list", '"red,blue"'),
        (second, "A1", "custom", "A1>0"),
    ):
        rule = DataValidation(type=kind, formula1=formula)
        if kind == "whole":
            rule.operator = "greaterThan"
        sheet.add_data_validation(rule)
        rule.add(target)
    first.conditional_formatting.add("A1:A20", CellIsRule(operator="greaterThan", formula=["3"]))
    second.conditional_formatting.add("A1", FormulaRule(formula=["A1>0"]))

    raw = BytesIO()
    book.save(raw)
    output = Path(__file__).resolve().parents[1] / "benchmarks" / "xlsx-rule-loss.openpyxl.xlsx"
    with ZipFile(raw) as source, ZipFile(output, "w") as target:
        for name in sorted(source.namelist()):
            data = source.read(name)
            if name == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]+", rb"\g<1>2026-11-10T00:00:00Z", data)
            target.writestr(ZipInfo(name, (2026, 11, 10, 0, 0, 0)), data, compress_type=ZIP_DEFLATED)

    with output.open("rb") as stream:
        restored = load_workbook(stream)
        assert [sheet["A1"].value for sheet in restored] == [5, 6]
        assert [[rule.type for rule in sheet.data_validations.dataValidation] for sheet in restored] == [["whole", "list"], ["custom"]]
        assert [len(sheet.conditional_formatting) for sheet in restored] == [1, 1]
        restored.close()
    print(output)


if __name__ == "__main__":
    main()

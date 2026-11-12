"""Create and independently inspect a synthetic defined-name XLSX fixture."""

import hashlib
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName


def main():
    if len(sys.argv) not in (2, 3) or (len(sys.argv) == 3 and sys.argv[2] not in ("--formulas", "--qualified", "--qualified-global")):
        raise SystemExit("Usage: python3 scripts/xlsx-defined-names-fixture.py NEW.xlsx [--formulas|--qualified|--qualified-global]")
    path = Path(sys.argv[1])
    formulas = len(sys.argv) == 3 and sys.argv[2] == "--formulas"
    qualified = len(sys.argv) == 3 and sys.argv[2] == "--qualified"
    qualified_global = len(sys.argv) == 3 and sys.argv[2] == "--qualified-global"
    report = path.with_suffix(path.suffix + ".json")
    if path.exists() or report.exists():
        raise SystemExit("Refusing to overwrite evidence")
    book = Workbook()
    sheet = book.active
    sheet.title = "Named data"
    for value in [11, 22, 33]:
        sheet.append([value])
    book.defined_names.add(DefinedName("GlobalValue", attr_text="'Named data'!$A$1"))
    book.defined_names.add(DefinedName("OtherValue", attr_text="'Named data'!$A$2"))
    sheet.defined_names.add(DefinedName("LocalValue", attr_text="'Named data'!$A$3"))
    if formulas:
        sheet["B1"] = "=GlobalValue*2"
        sheet["B2"] = "=LocalValue+OtherValue"
        sheet["B3"] = '="GlobalValue"'
        sheet["B4"] = "=SUM(A1:A3)"
    if qualified:
        other = book.create_sheet("Other")
        other["A1"] = "='Named data'!LocalValue+1"
        other["A2"] = "='Named data'!$A$1"
        other["A3"] = "='[outside.xlsx]Named data'!LocalValue"
    if qualified_global:
        other = book.create_sheet("Other")
        other["A1"] = "='Named data'!GlobalValue+1"
        other["A2"] = "='Named data'!$A$1"
        other["A3"] = '="GlobalValue"'
        other["A4"] = "='[outside.xlsx]Named data'!GlobalValue"
    with path.open("xb") as output:
        book.save(output)
    book.close()
    restored = load_workbook(path)
    global_names = sorted(restored.defined_names)
    local_names = sorted(restored["Named data"].defined_names)
    values = [restored["Named data"][f"A{i}"].value for i in range(1, 4)]
    formula_values = [restored["Named data"][f"B{i}"].value for i in range(1, 5)] if formulas else []
    qualified_values = [restored["Other"][f"A{i}"].value for i in range(1, 5 if qualified_global else 4)] if qualified or qualified_global else []
    restored.close()
    if global_names != ["GlobalValue", "OtherValue"] or local_names != ["LocalValue"] or values != [11, 22, 33]:
        raise ValueError("Independent readback disagrees")
    if formulas and formula_values != ["=GlobalValue*2", "=LocalValue+OtherValue", '=\"GlobalValue\"', "=SUM(A1:A3)"]:
        raise ValueError("Independent formula readback disagrees")
    if qualified and qualified_values != ["='Named data'!LocalValue+1", "='Named data'!$A$1", "='[outside.xlsx]Named data'!LocalValue"]:
        raise ValueError("Independent qualified formula readback disagrees")
    if qualified_global and qualified_values != ["='Named data'!GlobalValue+1", "='Named data'!$A$1", '="GlobalValue"', "='[outside.xlsx]Named data'!GlobalValue"]:
        raise ValueError("Independent qualified global formula readback disagrees")
    data = {"producer": "openpyxl", "version": openpyxl.__version__, "sha256": hashlib.sha256(path.read_bytes()).hexdigest(), "globalNames": global_names, "localNames": local_names, "values": values, "formulas": formula_values, "qualifiedFormulas": qualified_values}
    with report.open("x", encoding="utf-8") as output:
        json.dump(data, output, indent=2)
        output.write("\n")
    print(json.dumps(data))


if __name__ == "__main__":
    main()

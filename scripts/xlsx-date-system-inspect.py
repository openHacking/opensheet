"""Inspect date systems, formulas and cached results without recalculating them."""

import json
import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.datetime import MAC_EPOCH


def inspect(path):
    formula_book = load_workbook(path, data_only=False)
    cached_book = load_workbook(path, data_only=True)
    try:
        sheet = formula_book['Dates']
        result = {
            'system': '1904' if formula_book.epoch == MAC_EPOCH else '1900',
            'dates': [sheet.cell(row, 1).value.isoformat() for row in range(1, 4)],
            'formats': [sheet.cell(row, 1).number_format for row in range(1, 4)],
            'formula': sheet['B1'].value,
            'cached': cached_book['Dates']['B1'].value,
        }
    finally:
        formula_book.close()
        cached_book.close()
    with ZipFile(path) as archive:
        xml = ElementTree.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        raw = {}
        for cell in xml.findall('.//{*}c'):
            value = cell.find('{*}v')
            if value is not None and value.text is not None:
                raw[cell.get('r')] = value.text
        result['rawSerials'] = [raw.get(f'A{row}') for row in range(1, 4)]
        result['rawFormulaCache'] = raw.get('B1')
    return result


if __name__ == '__main__':
    if len(sys.argv) < 2:
        raise SystemExit('Usage: python3 scripts/xlsx-date-system-inspect.py FILE...')
    print(json.dumps({'reader': f'openpyxl {openpyxl.__version__}', 'files': [inspect(Path(path)) for path in sys.argv[1:]]}))

"""Create and independently read back synthetic 1900/1904 XLSX date fixtures."""

import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, to_excel


DATES = [datetime(1904, 1, 2), datetime(2024, 2, 29, 12, 30), datetime(2100, 3, 1)]
EXPECTED = {
    '1900': [1463, 45351.520833333336, 73110],
    '1904': [1, 43889.520833333336, 71648],
}


def main():
    if len(sys.argv) != 2:
        raise SystemExit('Usage: python3 scripts/xlsx-date-system-fixture.py NEW_DIRECTORY')
    destination = Path(sys.argv[1])
    destination.mkdir(parents=True, exist_ok=False)
    for system, epoch in [('1900', WINDOWS_EPOCH), ('1904', MAC_EPOCH)]:
        path = destination / f'dates-{system}.xlsx'
        book = Workbook()
        book.epoch = epoch
        sheet = book.active
        sheet.title = 'Dates'
        for row, value in enumerate(DATES, 1):
            cell = sheet.cell(row, 1, value)
            cell.number_format = 'yyyy-mm-dd hh:mm' if row == 2 else 'yyyy-mm-dd'
        sheet['B1'] = '=A2-A1'
        book.save(path)
        book.close()
        restored = load_workbook(path, data_only=False)
        values = [restored['Dates'].cell(row, 1).value.isoformat() for row in range(1, 4)]
        formats = [restored['Dates'].cell(row, 1).number_format for row in range(1, 4)]
        formula = restored['Dates']['B1'].value
        actual_epoch = '1904' if restored.epoch == MAC_EPOCH else '1900'
        restored.close()
        serials = [to_excel(value, epoch) for value in DATES]
        assert values == [value.isoformat() for value in DATES]
        assert formats == ['yyyy-mm-dd', 'yyyy-mm-dd hh:mm', 'yyyy-mm-dd']
        assert formula == '=A2-A1' and actual_epoch == system and serials == EXPECTED[system]
        with ZipFile(path) as archive:
            xml = archive.read('xl/workbook.xml').decode('utf-8')
            raw_epoch = 'date1904="1"' in xml
            assert raw_epoch == (system == '1904')
            worksheet = ElementTree.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            raw_values = {cell.get('r'): cell.find('{*}v').text for cell in worksheet.findall('.//{*}c') if cell.find('{*}v') is not None and cell.find('{*}v').text is not None}
            xlsx_serials = [float(raw_values[f'A{row}']) for row in range(1, 4)]
            assert all(abs(actual - expected) < 1e-9 for actual, expected in zip(xlsx_serials, serials))
        report = {
            'producer': 'openpyxl', 'version': openpyxl.__version__, 'system': system,
            'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'values': values,
            'serials': serials, 'xlsxSerials': xlsx_serials, 'formats': formats, 'formula': formula, 'rawDate1904': raw_epoch,
        }
        path.with_suffix('.xlsx.json').write_text(json.dumps(report, indent=2) + '\n')
        print(json.dumps(report))


if __name__ == '__main__':
    main()

"""Create or inspect a synthetic openpyxl workbook with one OOXML range hyperlink."""
import json
import sys
from datetime import datetime, timezone
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import openpyxl
from openpyxl import Workbook, load_workbook


def create(path):
    book = Workbook()
    book.properties.created = book.properties.modified = datetime(2026, 11, 21, tzinfo=timezone.utc)
    sheet = book.active
    sheet.title = 'Range'
    for value in ('one', 'two', 'three'):
        sheet.append([value])
    sheet['A1'].hyperlink = 'https://example.com/路径#片段'
    raw = BytesIO()
    book.save(raw)
    with ZipFile(raw) as source, ZipFile(path, 'x') as target:
        for name in sorted(source.namelist()):
            data = source.read(name)
            if name == 'xl/worksheets/sheet1.xml':
                old = b'ref="A1" r:id='
                assert data.count(old) == 1
                data = data.replace(old, b'ref="A1:A3" r:id=')
            target.writestr(ZipInfo(name, (2026, 11, 21, 0, 0, 0)), data, compress_type=ZIP_DEFLATED)


def inspect(path):
    book = load_workbook(path)
    sheet = book['Range']
    cells = []
    for row in range(1, 4):
        cell = sheet[f'A{row}']
        link = cell.hyperlink
        cells.append({'value': cell.value, 'target': link.target if link else None,
                      'location': link.location if link else None})
    book.close()
    with ZipFile(path) as archive:
        xml = archive.read('xl/worksheets/sheet1.xml').decode('utf-8')
        import re
        refs = re.findall(r'<hyperlink\s[^>]*?\bref="([^"]+)"', xml)
    return {'openpyxl': openpyxl.__version__, 'cells': cells, 'rawRefs': refs}


if __name__ == '__main__':
    if len(sys.argv) != 3 or sys.argv[1] not in ('create', 'inspect'):
        raise SystemExit('Usage: python3 scripts/xlsx-range-link-fixture.py create|inspect PATH')
    if sys.argv[1] == 'create':
        create(sys.argv[2])
    else:
        print(json.dumps(inspect(sys.argv[2]), ensure_ascii=False))

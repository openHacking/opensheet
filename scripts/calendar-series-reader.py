"""Compare anchored calendar sequences with LibreOffice EDATE recalculation.

Usage: python3 scripts/calendar-series-reader.py NEW_PREFIX
Creates source/saved XLSX files for both date systems and a JSON evidence report.
"""
import hashlib
import json
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, to_excel

CASES = [
    ('month_end', '2020-01-31', 1, 'month', 6),
    ('reverse_month_end', '2021-03-31', -1, 'month', 6),
    ('two_months', '2020-12-31', 2, 'month', 5),
    ('leap_year', '2020-02-29', 1, 'year', 5),
    ('reverse_year', '2024-02-29', -1, 'year', 5),
    ('century', '2096-02-29', 1, 'year', 6),
    ('stationary', '2000-05-31', 0, 'month', 4),
]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main(prefix):
    paths = {system: {kind: Path(f'{prefix}.{system}.{kind}.xlsx') for kind in ('source', 'saved')} for system in ('1900', '1904')}
    report_path = Path(f'{prefix}.json')
    if any(path.exists() for variants in paths.values() for path in variants.values()) or report_path.exists():
        raise SystemExit('Refusing to overwrite existing evidence')
    reader = subprocess.check_output(['soffice', '--version'], text=True, timeout=30).strip()
    results = []
    with tempfile.TemporaryDirectory(prefix='opensheet-calendar-reader-') as directory:
        temp = Path(directory)
        for system, variants in paths.items():
            book = Workbook()
            book.remove(book.active)
            book.epoch = CALENDAR_WINDOWS_1900 if system == '1900' else CALENDAR_MAC_1904
            for name, start, step, unit, count in CASES:
                sheet = book.create_sheet(name)
                serial = to_excel(datetime.fromisoformat(start), book.epoch)
                sheet['A1'] = 'start serial'
                sheet['B1'] = serial
                sheet['A2'] = 'step'
                sheet['B2'] = step
                for index in range(count):
                    sheet.cell(index + 4, 1, index)
                    sheet.cell(index + 4, 2, f'=EDATE($B$1,{index * step * (12 if unit == "year" else 1)})')
            book.save(variants['source'])
            log = subprocess.check_output(['soffice', f'-env:UserInstallation={(temp / "profile").as_uri()}', '--headless', '--convert-to', 'xlsx', '--outdir', str(temp), str(variants['source'].resolve())], text=True, timeout=60).strip()
            converted = temp / variants['source'].name
            variants['saved'].write_bytes(converted.read_bytes())
            converted.unlink()
            cached = load_workbook(variants['saved'], data_only=True)
            formulas = load_workbook(variants['saved'], data_only=False)
            if cached.epoch != book.epoch:
                raise RuntimeError('Reader changed date system')
            for name, start, step, unit, count in CASES:
                original_serial = book[name]['B1'].value
                actual = [cached[name].cell(index + 4, 2).value for index in range(count)]
                rewritten = [formulas[name].cell(index + 4, 2).value for index in range(count)]
                if not all(isinstance(value, (int, float)) for value in actual):
                    raise RuntimeError(f'Non-numeric result for {system}/{name}: {actual}')
                results.append({'system': system, 'case': name, 'date': start, 'start': original_serial,
                                'step': step, 'unit': unit, 'values': actual, 'readerFormulas': rewritten})
            cached.close()
            formulas.close()
            book.close()
    report = {'reader': reader, 'openpyxl': openpyxl.__version__, 'scope': 'Synthetic independently authored EDATE formulas on numeric serials, anchored to the original date, recalculated by one LibreOfficeDev build. This checks formula-derived dates, not native fill gestures, Excel, visual rendering or the 1900 fictitious leap day.',
              'cases': results,
              'artifacts': {f'{system}_{kind}': {'file': path.name, 'sha256': sha(path)} for system, variants in paths.items() for kind, path in variants.items()},
              'scriptSha256': sha(Path(__file__)), 'conversionLog': log}
    with report_path.open('x', encoding='utf8') as out:
        json.dump(report, out, ensure_ascii=False, indent=2)
        out.write('\n')
    print(f'{len(results)} sequences written to {report_path}')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    main(sys.argv[1])

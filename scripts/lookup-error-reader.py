"""Compare a hand-authored lookup corpus against a headless independent reader."""

from hashlib import sha256
import json
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory

import openpyxl
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / 'tests/fixtures/lookup-error-corpus.json'


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def main():
    if len(sys.argv) != 2 or not sys.argv[1].endswith('.json'):
        raise SystemExit('Supply a new report path ending in .json')
    report_path = Path(sys.argv[1]).resolve()
    source_path = report_path.with_suffix('.source.xlsx')
    saved_path = report_path.with_suffix('.saved.xlsx')
    if any(path.exists() for path in (report_path, source_path, saved_path)):
        raise SystemExit('Refusing to overwrite evidence')
    fixture = json.loads(FIXTURE.read_text())
    book = Workbook()
    sheet = book.active
    sheet.title = 'Lookup'
    for index, value in enumerate(fixture['inputs'], 1):
        if isinstance(value, dict):
            value = value['error'] if 'error' in value else '=' + value['formula']
        if value is not None:
            sheet.cell(index, 1, value)
    for index, value in enumerate(fixture['results'], 1):
        sheet.cell(index, 2, value)
    for index, case in enumerate(fixture['cases'], 1):
        sheet.cell(index, 4, '=' + case['formula'])
        sheet.cell(index, 5, case['formula'])
    version = subprocess.check_output(['soffice', '--version'], text=True, timeout=30).strip()
    with TemporaryDirectory() as directory:
        temp = Path(directory)
        source = temp / 'source' / 'lookup.xlsx'
        source.parent.mkdir()
        output = temp / 'saved'
        output.mkdir()
        book.save(source)
        subprocess.check_output(['soffice', f'-env:UserInstallation={(temp / "profile").as_uri()}', '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)], text=True, stderr=subprocess.STDOUT, timeout=60)
        saved = output / source.name
        with saved.open('rb') as stream:
            recalculated = load_workbook(stream, data_only=True)
            values = recalculated['Lookup']
            observations = []
            for index, case in enumerate(fixture['cases'], 1):
                cell = values.cell(index, 4)
                actual = {'error': cell.value} if cell.data_type == 'e' else cell.value
                observations.append({**case, 'reader': actual, 'agrees': actual == case['expected']})
            recalculated.close()
        shutil.copyfile(source, source_path)
        shutil.copyfile(saved, saved_path)
    report = {
        'reader': version,
        'inspector': f'openpyxl {openpyxl.__version__}',
        'scope': 'Synthetic exact MATCH and VLOOKUP cases with lookup errors and blank cells. One LibreOfficeDev alpha build, headless recalculation; no Excel or stable LibreOffice claim.',
        'fixture': {'file': str(FIXTURE.relative_to(ROOT)), 'sha256': digest(FIXTURE)},
        'artifacts': {'source': {'file': source_path.name, 'sha256': digest(source_path)}, 'saved': {'file': saved_path.name, 'sha256': digest(saved_path)}},
        'observations': observations,
    }
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n')
    print(json.dumps({'reader': version, 'agreements': sum(item['agrees'] for item in observations), 'observations': observations}, indent=2))


if __name__ == '__main__':
    main()

"""Create and read back a synthetic XLSX containing Unicode defined names."""
import hashlib
import json
import sys
from pathlib import Path
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

if len(sys.argv) != 2:
    raise SystemExit("Usage: python3 scripts/xlsx-unicode-names-fixture.py NEW.xlsx")
path = Path(sys.argv[1])
report = path.with_suffix(path.suffix + '.json')
if path.exists() or report.exists():
    raise SystemExit("Refusing to overwrite evidence")
book = Workbook()
data = book.active
data.title = '数据'
data['A1'] = 12
data['A2'] = 7
book.defined_names.add(DefinedName('预算', attr_text="'数据'!$A$1"))
data.defined_names.add(DefinedName('总计', attr_text="'数据'!$A$2"))
data['B1'] = '=预算+1'
data['B2'] = '=总计+1'
data['B3'] = '="预算"'
other = book.create_sheet('Other')
other['A1'] = "='数据'!总计+1"
other['A2'] = "='数据'!预算+1"
other['A3'] = "='数据'!$A$1"
other['A4'] = "='[external.xlsx]数据'!预算"
with path.open('xb') as out:
    book.save(out)
book.close()
restored = load_workbook(path)
observed = {'global':sorted(restored.defined_names), 'local':sorted(restored['数据'].defined_names),
            'dataFormulas':[restored['数据'][f'B{i}'].value for i in range(1,4)],
            'otherFormulas':[restored['Other'][f'A{i}'].value for i in range(1,5)]}
restored.close()
assert observed == {'global':['预算'], 'local':['总计'],
                    'dataFormulas':['=预算+1','=总计+1','="预算"'],
                    'otherFormulas':["='数据'!总计+1","='数据'!预算+1","='数据'!$A$1","='[external.xlsx]数据'!预算"]}
result = {'producer':'openpyxl','version':openpyxl.__version__,'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),**observed}
with report.open('x', encoding='utf8') as out:
    json.dump(result,out,ensure_ascii=False,indent=2)
    out.write('\n')
print(json.dumps(result,ensure_ascii=False))

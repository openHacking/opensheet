"""Create and independently inspect an XLSX with backslash-prefixed names."""
import hashlib
import json
import sys
from pathlib import Path
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

if len(sys.argv) != 2:
    raise SystemExit('Usage: python3 scripts/xlsx-backslash-names-fixture.py NEW.xlsx')
path = Path(sys.argv[1])
report = path.with_suffix(path.suffix + '.json')
if path.exists() or report.exists():
    raise SystemExit('Refusing to overwrite evidence')
book = Workbook()
data = book.active
data.title = 'Data'
data['A1'] = 11
data['A2'] = 22
book.defined_names.add(DefinedName('\\Budget', attr_text="'Data'!$A$1"))
data.defined_names.add(DefinedName('\\本地', attr_text="'Data'!$A$2"))
data['B1'] = '=\\Budget+1'
data['B2'] = '=\\本地+1'
data['B3'] = '="\\Budget"'
other = book.create_sheet('Other')
other['A1'] = "='Data'!\\Budget+1"
other['A2'] = "='Data'!\\本地+1"
other['A3'] = "='Data'!$A$1"
other['A4'] = "='[external.xlsx]Data'!\\Budget"
with path.open('xb') as out:
    book.save(out)
book.close()
restored = load_workbook(path)
observed = {'global':sorted(restored.defined_names),'local':sorted(restored['Data'].defined_names),
            'dataFormulas':[restored['Data'][f'B{i}'].value for i in range(1,4)],
            'otherFormulas':[restored['Other'][f'A{i}'].value for i in range(1,5)]}
restored.close()
assert observed == {'global':['\\Budget'],'local':['\\本地'],
                    'dataFormulas':['=\\Budget+1','=\\本地+1','="\\Budget"'],
                    'otherFormulas':["='Data'!\\Budget+1","='Data'!\\本地+1","='Data'!$A$1","='[external.xlsx]Data'!\\Budget"]}
result = {'producer':'openpyxl','version':openpyxl.__version__,'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),**observed}
with report.open('x',encoding='utf8') as out:
    json.dump(result,out,ensure_ascii=False,indent=2)
    out.write('\n')
print(json.dumps(result,ensure_ascii=False))

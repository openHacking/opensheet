"""Create an independent unfrozen worksheet with a selected tab and active cell."""
import argparse
import json
from pathlib import Path
import openpyxl
from openpyxl.worksheet.views import Selection

parser = argparse.ArgumentParser()
parser.add_argument('output')
args = parser.parse_args()
if Path(args.output).exists():
    raise FileExistsError('Refusing to overwrite fixture')
book = openpyxl.Workbook()
sheet = book.active
sheet.title = 'View'
sheet['A1'] = 42
sheet.sheet_view.tabSelected = True
sheet.sheet_view.selection = [Selection(activeCell='D8', sqref='D8')]
book.save(args.output)
saved = openpyxl.load_workbook(args.output)
view = saved['View'].sheet_view
print(json.dumps({'openpyxl': openpyxl.__version__, 'tabSelected': view.tabSelected,
                  'activeCell': view.selection[0].activeCell, 'sqref': view.selection[0].sqref}))

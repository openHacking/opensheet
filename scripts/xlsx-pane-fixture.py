"""Create a synthetic XLSX with independent frozen and split pane metadata."""
import argparse
import json
import openpyxl
from openpyxl.worksheet.views import Pane, Selection

parser = argparse.ArgumentParser()
parser.add_argument('output')
args = parser.parse_args()
book = openpyxl.Workbook()
frozen = book.active
frozen.title = 'Frozen'
frozen['A1'] = 7
frozen.freeze_panes = 'C4'
split = book.create_sheet('Split')
split['A1'] = 9
split.sheet_view.pane = Pane(xSplit=1200, ySplit=900, topLeftCell='B2', activePane='bottomRight', state='split')
split.sheet_view.selection = [Selection(pane='bottomRight', activeCell='B2', sqref='B2')]
book.save(args.output)
check = openpyxl.load_workbook(args.output)
print(json.dumps({'openpyxl': openpyxl.__version__, 'frozen': check['Frozen'].freeze_panes,
                  'split': check['Split'].sheet_view.pane.state}))

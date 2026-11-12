"""Generate independent XLSX frozen-pane variants for import verification."""
import argparse
import json
import openpyxl

parser = argparse.ArgumentParser()
parser.add_argument('output')
args = parser.parse_args()
book = openpyxl.Workbook()
for index, (name, pane) in enumerate([('Rows', 'A5'), ('Columns', 'D1'), ('Both', 'C4')]):
    sheet = book.active if index == 0 else book.create_sheet()
    sheet.title = name
    sheet['A1'] = index + 1
    sheet.freeze_panes = pane
book.save(args.output)
readback = openpyxl.load_workbook(args.output)
print(json.dumps({'openpyxl': openpyxl.__version__, 'panes': [readback[name].freeze_panes for name in readback.sheetnames]}))

"""Create an independent XLSX dimensions fixture, or inspect an OpenSheet export."""
import argparse
import json
import openpyxl

parser = argparse.ArgumentParser()
parser.add_argument('mode', choices=['create', 'inspect'])
parser.add_argument('path')
args = parser.parse_args()
if args.mode == 'create':
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = 'Dimensions'
    sheet['A1'] = 1
    sheet.row_dimensions[101].height = 31.5
    sheet.row_dimensions[101].hidden = True
    sheet.column_dimensions['U'].width = 20
    sheet.column_dimensions['U'].hidden = True
    book.save(args.path)
book = openpyxl.load_workbook(args.path)
sheet = book['Dimensions']
print(json.dumps({'openpyxl': openpyxl.__version__, 'row_points': sheet.row_dimensions[101].height,
                  'row_hidden': sheet.row_dimensions[101].hidden,
                  'column_width': sheet.column_dimensions['U'].width,
                  'column_hidden': sheet.column_dimensions['U'].hidden}))

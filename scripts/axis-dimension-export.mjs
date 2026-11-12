import { writeFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';

const output = process.argv[2];
if (!output) throw new Error('Pass a new output path');
const file = createWorkbookFile({ sheets: [{ name: 'Dimensions', rows: 120, columns: 26 }] });
file.sheets[0].rows.meta = { 100: { size: 42, hidden: true } };
file.sheets[0].columns.meta = { 20: { size: 140, hidden: true } };
const result = toSheetJS(file);
const bytes = XLSX.write(result.workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
writeFileSync(output, bytes, { flag: 'wx' });
console.log(JSON.stringify(result.report));

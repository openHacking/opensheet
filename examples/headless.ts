import { createWorkbook } from '@opensheetjs/core';
import { exportRange } from '@opensheetjs/formats';
const workbook = createWorkbook();
const sheet = workbook.getSheets()[0];
workbook.transaction({ label: 'Example' }, () => {
  sheet.range('A1:B3').setValues([
    ['Month', 'Revenue'],
    ['Jan', 100],
    ['Feb', 200],
  ]);
  sheet.range('B4').setFormulas([['SUM(B2:B3)']]);
});
const result = exportRange(workbook.toJSON(), { sheetId: sheet.id, format: 'markdown' });
console.log(result.text);
workbook.dispose();

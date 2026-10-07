import { createWorkbook } from '@opensheetjs/core';
import { exportRange } from '@opensheetjs/formats';
const workbook = await createWorkbook();
const sheet = workbook.getSheets()[0];
await workbook.transaction({ label: 'Example' }, async (workbook) => {
  await workbook
    .getSheetById(sheet.id)!
    .range('A1:B3')
    .setValues([
      ['Month', 'Revenue'],
      ['Jan', 100],
      ['Feb', 200],
    ]);
  await workbook
    .getSheetById(sheet.id)!
    .range('B4')
    .setFormulas([['SUM(B2:B3)']]);
});
const result = exportRange(await workbook.toJSON(), { sheetId: sheet.id, format: 'markdown' });
console.log(result.text);
workbook.dispose();

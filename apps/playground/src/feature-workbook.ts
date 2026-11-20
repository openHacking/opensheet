import { createWorkbook, type WorkbookFile } from 'opensheet';

export async function createFeatureWorkbook(): Promise<WorkbookFile> {
  const book = await createWorkbook({
    sheets: ['Orders', 'Rates', 'Practice'].map((name) => ({ name, rows: 100, columns: 12 })),
  });
  try {
    const [orders, rates, practice] = book.getSheets();
    await book.transaction({ label: 'All Features template' }, (transaction) => {
      const orderSheet = transaction.getSheetById(orders.id)!;
      const rateSheet = transaction.getSheetById(rates.id)!;
      const practiceSheet = transaction.getSheetById(practice.id)!;
      orderSheet.range('A1:G7').setValues([
        ['Product', 'Region', 'Units', 'Unit price', 'Total', 'Status', 'Paid'],
        ['Notebook', 'West', 12, 8.5, null, 'Shipped', true],
        ['Pencil set', 'East', 6, 4.75, null, 'Pending', false],
        ['Desk pad', 'West', 3, 18, null, 'Shipped', true],
        ['Notebook', 'East', 9, 8.5, null, 'Pending', false],
        ['Folder', 'North', 15, 3.2, null, 'Shipped', true],
        ['Pencil set', 'North', 8, 4.75, null, 'Pending', false],
      ]);
      orderSheet.range('E2:E7').setFormulas(
        Array.from({ length: 6 }, (_, index) => {
          const row = index + 2;
          return [`ROUND(C${row}*D${row}*(1-Rates!$B$2),2)`];
        }),
      );
      orderSheet.range('D2:E7').setStyle({ numberFormat: '$#,##0.00' });
      orderSheet.range('A2').setNote('Bulk discount applies to this order. Change it on Rates.');
      orderSheet.range('A2').setLink({
        target: 'https://github.com/openHacking/opensheet',
        tooltip: 'OpenSheet on GitHub',
      });
      orderSheet.setFreeze(1, 1);
      orderSheet.setColumnWidth(0, 126);

      rateSheet.range('A1:B3').setValues([
        ['Setting', 'Value'],
        ['Discount', 0.1],
        ['Order total', null],
      ]);
      rateSheet.range('B2').setStyle({ numberFormat: '0.0%' });
      rateSheet.range('B3').setFormulas([['SUM(Orders!E2:E7)']]);
      rateSheet.range('B3').setStyle({ numberFormat: '$#,##0.00' });
      rateSheet.setColumnWidth(0, 126);

      practiceSheet.range('A1:D2').setValues([
        ['Series', 'Fill formula', 'Paste here', 'Date'],
        [1, null, null, 46346],
      ]);
      practiceSheet.range('A3').setValues([[2]]);
      practiceSheet.range('B2').setFormulas([['A2*10']]);
      practiceSheet.range('D2').setStyle({ numberFormat: 'yyyy-mm-dd' });
      practiceSheet.range('A8').setValues([['Merged heading']]);
      practiceSheet.range('A8:C8').merge();
      practiceSheet.range('A8:C8').setStyle({ align: 'center', bold: true, border: true });
      practiceSheet.range('A9').setValues([['Wrap text here, then resize the row or column.']]);
      practiceSheet.setColumnWidth(0, 126);

      for (const [sheet, header] of [
        [orderSheet, 'A1:G1'],
        [rateSheet, 'A1:B1'],
        [practiceSheet, 'A1:D1'],
      ] as const) {
        sheet.range(header).setStyle({ bold: true, color: '#235b42', background: '#eef7f1' });
        sheet.setRowHeight(0, 36);
      }
    });
    return await book.toJSON();
  } finally {
    try {
      await book.deleteStorage();
    } finally {
      book.dispose();
    }
  }
}

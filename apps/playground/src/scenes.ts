import { createWorkbook, type WorkbookSnapshot } from 'opensheet';
export type SceneId = 'budget' | 'sales' | 'planner' | 'code' | 'performance';
export const scenes: Record<
  SceneId,
  {
    title: string;
    label: string;
    description: string;
    icon: string;
  }
> = {
  performance: {
    title: 'Performance Lab',
    label: '05 / SCALE',
    description: 'Find the largest editable workbook within a bounded local storage budget.',
    icon: '↗',
  },
  budget: {
    title: 'Budget Simulator',
    label: '01 / PLAN',
    description: 'Turn a what-if into a clear plan. Edit prices and watch the numbers move.',
    icon: '↗',
  },
  sales: {
    title: 'Sales Dashboard',
    label: '02 / MEASURE',
    description: 'From monthly numbers to a live picture of your business.',
    icon: '▥',
  },
  planner: {
    title: 'Project Planner',
    label: '03 / SHIP',
    description: 'A spreadsheet with a sense of time. Plan the next launch.',
    icon: '→',
  },
  code: {
    title: 'Table to Code',
    label: '04 / CREATE',
    description: 'One table. Five formats. Ready for your next document.',
    icon: '〈〉',
  },
};
export async function createScene(id: Exclude<SceneId, 'budget'>): Promise<WorkbookSnapshot> {
  const book = await createWorkbook({
    sheets: [{ name: scenes[id].title, rows: 100, columns: 12 }],
  });
  const sheet = book.getSheets()[0];
  await book.transaction({ label: 'Sample workbook' }, (book) => {
    if (id === 'sales') {
      book
        .getSheetById(sheet.id)!
        .range('A1:E8')
        .setValues([
          ['Month', 'Revenue', 'Target', 'Growth', 'Attainment'],
          ['January', 18000, 20000, null, null],
          ['February', 23500, 22000, null, null],
          ['March', 27800, 25000, null, null],
          ['April', 31200, 28000, null, null],
          ['May', 36400, 32000, null, null],
          ['June', 42500, 38000, null, null],
          ['TOTAL', null, null, null, null],
        ]);
      book
        .getSheetById(sheet.id)!
        .range('B8:C8')
        .setFormulas([['SUM(B2:B7)', 'SUM(C2:C7)']]);
      book
        .getSheetById(sheet.id)!
        .range('D3:D7')
        .setFormulas(Array.from({ length: 5 }, (_, i) => [`B${i + 3}/B${i + 2}-1`]));
      book
        .getSheetById(sheet.id)!
        .range('E2:E7')
        .setFormulas(Array.from({ length: 6 }, (_, i) => [`B${i + 2}/C${i + 2}`]));
      book.getSheetById(sheet.id)!.range('B2:C8').setStyle({ numberFormat: '$#,##0' });
      book.getSheetById(sheet.id)!.range('D2:E7').setStyle({ numberFormat: '0.0%' });
      book.getSheetById(sheet.id)!.range('A8:E8').setStyle({ bold: true, background: '#e4f6ed' });
    } else if (id === 'planner') {
      book
        .getSheetById(sheet.id)!
        .range('A1:E7')
        .setValues([
          ['Task', 'Owner', 'Start date', 'Duration (days)', 'Progress'],
          ['Discovery', 'Olivia', '2026-10-01', 5, 1],
          ['Design system', 'Sam', '2026-10-04', 8, 0.75],
          ['Editor prototype', 'Alex', '2026-10-09', 10, 0.4],
          ['Integration tests', 'Jordan', '2026-10-16', 7, 0.15],
          ['Documentation', 'Sam', '2026-10-19', 6, 0.1],
          ['Launch', 'Olivia', '2026-10-27', 3, 0],
        ]);
      book.getSheetById(sheet.id)!.range('E2:E7').setStyle({ numberFormat: '0%' });
    } else {
      book
        .getSheetById(sheet.id)!
        .range('A1:D6')
        .setValues([
          ['Model', 'Accuracy', 'Latency (ms)', 'Parameters (M)'],
          ['Baseline', 0.872, 42, 12],
          ['Compact', 0.891, 28, 8],
          ['Balanced', 0.923, 35, 16],
          ['Extended', 0.948, 57, 32],
          ['AVERAGE', null, null, null],
        ]);
      book
        .getSheetById(sheet.id)!
        .range('B6:D6')
        .setFormulas([['AVERAGE(B2:B5)', 'AVERAGE(C2:C5)', 'AVERAGE(D2:D5)']]);
      book.getSheetById(sheet.id)!.range('B2:B6').setStyle({ numberFormat: '0.0%' });
      book.getSheetById(sheet.id)!.range('A6:D6').setStyle({ bold: true, background: '#e4f6ed' });
    }
    book
      .getSheetById(sheet.id)!
      .range(id === 'code' ? 'A1:D1' : 'A1:E1')
      .setStyle({ bold: true, color: '#235b42', background: '#eef7f1' });
    book.getSheetById(sheet.id)!.setColumnWidth(0, 210);
    for (let column = 1; column < 5; column++)
      book.getSheetById(sheet.id)!.setColumnWidth(column, 155);
    book.getSheetById(sheet.id)!.setRowHeight(0, 36);
    book.getSheetById(sheet.id)!.setFreeze(1);
  });
  const snapshot = await book.toJSON();
  await book.deleteStorage();
  book.dispose();
  return snapshot;
}

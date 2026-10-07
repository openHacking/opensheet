import { createWorkbook, type WorkbookSnapshot } from 'opensheet';
export function sample(): WorkbookSnapshot {
  const b = createWorkbook({
    sheets: [
      { name: 'Launch budget', rows: 100, columns: 12 },
      { name: 'Notes', rows: 100, columns: 12 },
    ],
  });
  const s = b.getSheets()[0];
  b.transaction({ label: 'Example' }, () => {
    s.range('A1:G11').setValues([
      ['Item', 'Category', 'Owner', 'Quantity', 'Unit cost', 'Total', 'Status'],
      ['Brand identity', 'Design', 'Olivia', 1, 2400, null, 'Complete'],
      ['Landing page', 'Development', 'Alex', 1, 3800, null, 'In progress'],
      ['Product photography', 'Content', 'Sam', 2, 450, null, 'Complete'],
      ['Launch video', 'Content', 'Jordan', 1, 1800, null, 'In progress'],
      ['Social campaign', 'Marketing', 'Olivia', 4, 350, null, 'Planned'],
      ['Email platform', 'Software', 'Alex', 3, 49, null, 'Complete'],
      ['Press kit', 'Content', 'Sam', 1, 650, null, 'Planned'],
      ['Community event', 'Marketing', 'Jordan', 1, 1200, null, 'Planned'],
      ['Contingency', 'Operations', 'Alex', 1, 800, null, 'Reserved'],
      ['TOTAL', null, null, null, null, null, null],
    ]);
    s.range('F2:F10').setFormulas(Array.from({ length: 9 }, (_, i) => [`D${i + 2}*E${i + 2}`]));
    s.range('F11').setFormulas([['SUM(F2:F10)']]);
    s.range('A1:G1').setStyle({ bold: true, background: '#f0f5f1', color: '#476451' });
    s.range('E2:F11').setStyle({ numberFormat: '$#,##0.00' });
    s.range('A11:G11').setStyle({ bold: true, background: '#eef6f0' });
    s.range('G2:G10').setStyle({ color: '#248363', fontSize: 12 });
    s.setColumnWidth(0, 205);
    s.setColumnWidth(1, 142);
    s.setColumnWidth(2, 112);
    s.setColumnWidth(3, 92);
    s.setColumnWidth(4, 118);
    s.setColumnWidth(5, 130);
    s.setColumnWidth(6, 145);
    s.setRowHeight(0, 36);
    s.setFreeze(1);
    b.getSheets()[1]
      .range('A1:B5')
      .setValues([
        ['OpenSheet', 'Quick notes'],
        ['Edit', 'Double-click a cell or press Enter.'],
        ['Formula', '=SUM(F2:F10)'],
        ['Privacy', 'All processing happens on this device.'],
        ['Compatibility', 'XLSX exports preserve supported data, not every Excel feature.'],
      ]);
    b.getSheets()[1].setColumnWidth(1, 540);
  });
  const json = b.toJSON();
  b.dispose();
  return json;
}

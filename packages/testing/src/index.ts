import { createWorkbook, type Primitive, type WorkbookSnapshot, Workbook } from '@opensheetjs/core';
export function workbookFromRows(rows: Primitive[][], name = 'Sheet1') {
  const book = createWorkbook({
    sheets: [
      {
        name,
        rows: Math.max(100, rows.length),
        columns: Math.max(26, ...rows.map((r) => r.length)),
      },
    ],
  });
  if (rows.length)
    book
      .getSheets()[0]
      .range({ startRow: 0, startColumn: 0, endRow: rows.length, endColumn: rows[0].length })
      .setValues(rows);
  return book;
}
export function normalizedSnapshot(snapshot: WorkbookSnapshot) {
  const copy = structuredClone(snapshot);
  copy.revision = 0;
  return copy;
}
export function roundTrip(snapshot: WorkbookSnapshot) {
  const restored = new Workbook(JSON.parse(JSON.stringify(snapshot)));
  const result = restored.toJSON();
  restored.dispose();
  return result;
}

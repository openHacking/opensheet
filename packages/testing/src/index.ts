import { createWorkbook, type Primitive, type WorkbookFile, Workbook } from '@opensheetjs/core';
export async function workbookFromRows(rows: Primitive[][], name = 'Sheet1') {
  const book = await createWorkbook({
    sheets: [
      {
        name,
        rows: Math.max(100, rows.length),
        columns: Math.max(26, ...rows.map((r) => r.length)),
      },
    ],
  });
  if (rows.length)
    await book
      .getSheets()[0]
      .range({ startRow: 0, startColumn: 0, endRow: rows.length, endColumn: rows[0].length })
      .setValues(rows);
  return book;
}
export function normalizedSnapshot(snapshot: WorkbookFile) {
  const copy = structuredClone(snapshot);
  copy.revision = 0;
  return copy;
}
export async function roundTrip(snapshot: WorkbookFile) {
  const restored = await new Workbook(JSON.parse(JSON.stringify(snapshot))).ready();
  const result = await restored.toJSON();
  restored.dispose();
  return result;
}

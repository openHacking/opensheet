import { expect, it } from 'vitest';
import { createWorkbook, Workbook } from '@opensheetjs/core';
import { createSnapshot } from '../packages/core/src/model.js';
import { encodeWorkbookFile } from '../packages/core/src/file-codec.js';

it('finds visible region edges using inputs and keeps history intact', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 12, columns: 5 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A7').setValues([[1], [2], [null], [3], [4], [null], [5]]);
    await sheet.range('B1').setInput('=IF(TRUE,"",1)');
    await sheet.setRowHidden(1, true);
    const jump = async (row: number, direction: -1 | 1) =>
      (await book.navigateData(sheet.id, row, 0, 'row', direction)).row;
    expect(await jump(0, 1)).toBe(3);
    expect(await jump(3, 1)).toBe(4);
    expect(await jump(4, 1)).toBe(6);
    expect(await jump(6, 1)).toBe(11);
    expect(await jump(6, -1)).toBe(4);
    expect(await jump(4, -1)).toBe(3);
    expect(await jump(3, -1)).toBe(0);
    expect((await book.navigateData(sheet.id, 0, 0, 'column', 1)).column).toBe(1);
    await sheet.range('A8:A9').merge();
    await sheet.range('A8').setValues([[8]]);
    expect(await jump(6, 1)).toBe(8);
    expect(await jump(7, 1)).toBe(11);
    expect(await jump(8, -1)).toBe(6);
    await book.undo();
    expect(await jump(6, 1)).toBe(11);
    await book.redo();
    expect(await jump(6, 1)).toBe(8);
    await expect(book.navigateData(sheet.id, -1, 0, 'row', 1)).rejects.toMatchObject({
      code: 'INVALID_ARGUMENT',
    });
    const abort = new AbortController();
    abort.abort();
    await expect(
      book.navigateData(sheet.id, 0, 0, 'row', 1, { signal: abort.signal }),
    ).rejects.toMatchObject({ code: 'ABORTED' });
  } finally {
    await book.deleteStorage();
  }
});

it('matches an independent visible-position model across gaps and both directions', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 64, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    const filled = Array.from({ length: 64 }, (_, i) => (i * 17 + 3) % 11 < 5);
    const hidden = new Set([2, 3, 12, 25, 26, 27, 46, 62]);
    await sheet.range('A1:A64').setValues(filled.map((value, i) => [value ? i : null]));
    for (const row of hidden) await sheet.setRowHidden(row, true);
    const visible = filled.map((_, i) => i).filter((i) => !hidden.has(i));
    const model = (origin: number, direction: -1 | 1) => {
      let cursor = visible.indexOf(origin),
        next = cursor + direction;
      if (next < 0 || next >= visible.length) return origin;
      if (filled[origin] && filled[visible[next]]) {
        while (next >= 0 && next < visible.length && filled[visible[next]]) {
          cursor = next;
          next += direction;
        }
        return visible[cursor];
      }
      while (next >= 0 && next < visible.length) {
        cursor = next;
        if (filled[visible[next]]) break;
        next += direction;
      }
      return visible[cursor];
    };
    for (const origin of visible)
      for (const direction of [-1, 1] as const) {
        expect((await book.navigateData(sheet.id, origin, 0, 'row', direction)).row).toBe(
          model(origin, direction),
        );
      }
  } finally {
    await book.deleteStorage();
  }
});

it('skips filtered and hidden positions and follows restored row identities', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 6, columns: 4 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B6').setValues([
      [1, 'yes'],
      [2, 'no'],
      [3, 'yes'],
      [4, 'yes'],
      [5, 'no'],
      [6, 'yes'],
    ]);
    await sheet.setFilter({
      mode: 'all',
      headerRows: 0,
      criteria: [{ column: 1, operator: 'eq', value: 'yes' }],
    });
    await sheet.setRowHidden(2, true);
    expect(book.filteredRows(sheet.id)).toEqual([1, 4]);
    expect((await book.navigateData(sheet.id, 0, 0, 'row', 1)).row).toBe(5);
    await sheet.setFilter(null);
    expect((await book.navigateData(sheet.id, 0, 0, 'row', 1)).row).toBe(5);
    await sheet.insertRows(1);
    expect((await book.navigateData(sheet.id, 0, 0, 'row', 1)).row).toBe(2);
    await book.undo();
    expect((await book.navigateData(sheet.id, 0, 0, 'row', 1)).row).toBe(5);
    await sheet.range('A1:D1').setValues([[1, 2, 3, 4]]);
    await sheet.setColumnHidden(1, true);
    expect((await book.navigateData(sheet.id, 0, 0, 'column', 1)).column).toBe(3);
  } finally {
    await book.deleteStorage();
  }
});

it('skips zero-sized imported rows and columns', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 4, columns: 4 }] });
  const sheet = snapshot.sheets[0];
  sheet.rows['1'] = { size: 0 };
  sheet.columns['1'] = { size: 0 };
  for (let i = 0; i < 4; i++) {
    sheet.cells[`${i}:0`] = { rowId: i, columnId: 0, input: { type: 'number', value: i } };
    sheet.cells[`0:${i}`] = { rowId: 0, columnId: i, input: { type: 'number', value: i } };
  }
  const book = await new Workbook(encodeWorkbookFile(snapshot)).ready();
  try {
    expect((await book.navigateData(sheet.id, 0, 0, 'row', 1)).row).toBe(3);
    expect((await book.navigateData(sheet.id, 0, 0, 'column', 1)).column).toBe(3);
  } finally {
    await book.deleteStorage();
  }
});

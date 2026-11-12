import { describe, expect, it } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import { numericSeries } from '../packages/core/src/numeric-series.js';

it('validates paged fill targets and composes fill with other transaction commands', async () => {
  const book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A3').setValues([[1], [2], [3]]);
    const revision = book.revision;
    await expect(
      book.execute({
        type: 'core.cells.fillDown',
        payload: {
          sheetId: 'missing-sheet',
          range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
        },
      }),
    ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1:A3').getValues()).toEqual([[1], [2], [3]]);
    await book.transaction({}, async (tx) => {
      const target = tx.getSheetById(sheet.id)!;
      await target.range('A1').setValues([[9]]);
      await target.range('A1:A3').fillDown();
      await target.range('A3').setValues([[12]]);
    });
    expect(await sheet.range('A1:A3').getValues()).toEqual([[9], [9], [12]]);
    await book.undo();
    expect(await sheet.range('A1:A3').getValues()).toEqual([[1], [2], [3]]);
    await sheet.range('A2:B2').merge();
    const mergedRevision = book.revision;
    await expect(sheet.range('A1:A3').fillDown()).rejects.toMatchObject({ code: 'INVALID_RANGE' });
    expect(book.revision).toBe(mergedRevision);
    expect(await sheet.range('A1:A3').getValues()).toEqual([[1], [2], [3]]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

describe('numeric fill series', () => {
  it('uses decimal arithmetic for fractional, descending, constant and extreme inputs', () => {
    expect(numericSeries(0.1, 0.1, 4)).toEqual([0.1, 0.2, 0.3, 0.4]);
    expect(numericSeries(0.3, -0.1, 5)).toEqual([0.3, 0.2, 0.1, 0, -0.1]);
    expect(numericSeries(-2, 0, 3)).toEqual([-2, -2, -2]);
    expect(numericSeries(5e-324, 5e-324, 3)).toEqual([5e-324, 1e-323, 1.5e-323]);
    expect(numericSeries(1e308, -1e308, 3)).toEqual([1e308, 0, -1e308]);
    expect(numericSeries(-0, -0, 1)).toEqual([0]);
    expect(() => numericSeries(1e308, 1e308, 2)).toThrow();
  });

  it('fills either axis, preserves annotations and style, recalculates and restores history', async () => {
    let book = await createWorkbook();
    const target = await createWorkbook();
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('B2').setFormulas([['1+2']]);
      await sheet.range('D1').setFormulas([['B2*10']]);
      await sheet.range('B2').setStyle({ bold: true });
      await sheet.range('B2').setNote('Keep me');
      await sheet.range('B2').setLink({ target: 'https://example.com' });
      expect(await sheet.range('D1').getValues()).toEqual([[30]]);
      await sheet.range('B2:C5').fillSeries({ start: 0.1, step: 0.1 });
      const expected = [
        [0.1, 0.1],
        [0.2, 0.2],
        [0.3, 0.3],
        [0.4, 0.4],
      ];
      expect(await sheet.range('B2:C5').getValues()).toEqual(expected);
      expect(await sheet.range('D1').getValues()).toEqual([[1]]);
      expect(await sheet.range('B2').getFormulas()).toEqual([[null]]);
      expect(await sheet.range('B2').getNotes()).toEqual([['Keep me']]);
      expect(await sheet.range('B2').getLinks()).toEqual([[{ target: 'https://example.com' }]]);
      const record = await book.getCell(sheet.id, 1, 1);
      expect(book.getStyle(record?.styleId).bold).toBe(true);
      await book.undo();
      expect(await sheet.range('D1').getValues()).toEqual([[30]]);
      await book.redo();
      await target.importJSON(await book.toJSON());
      expect(await target.getSheets()[0].range('B2:C5').getValues()).toEqual(expected);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      sheet = book.getSheets()[0];
      expect(await sheet.range('B2:C5').getValues()).toEqual(expected);
      await sheet.range('B2:D3').fillSeries({ start: 3, step: -2, axis: 'column' });
      expect(await sheet.range('B2:D3').getValues()).toEqual([
        [3, 1, -1],
        [3, 1, -1],
      ]);
      await sheet.range('A1:A3').fillSeries({ start: 7 });
      expect(await sheet.range('A1:A3').getValues()).toEqual([[7], [8], [9]]);
    } finally {
      await book.deleteStorage();
      await book.close();
      await target.deleteStorage();
      await target.close();
    }
  });

  it('rejects overflow, invalid options and merged targets without publishing any partial write', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A3').setValues([[10], [20], [30]]);
      const revision = book.revision;
      for (const options of [
        { start: 1e308, step: 1e308 },
        { start: NaN },
        { start: Infinity },
        { start: 1, step: -Infinity },
        { start: 1, axis: 'diagonal' },
        { start: 1, extra: true },
      ]) {
        await expect(
          book.execute({
            type: 'core.cells.fillSeries',
            payload: {
              sheetId: sheet.id,
              range: sheet.range('A1:A3').bounds,
              ...options,
            },
          }),
        ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
        expect(book.revision).toBe(revision);
        expect(await sheet.range('A1:A3').getValues()).toEqual([[10], [20], [30]]);
      }
      await sheet.range('B2:C2').merge();
      const mergedRevision = book.revision;
      await expect(sheet.range('A1:C3').fillSeries({ start: 1 })).rejects.toMatchObject({
        code: 'INVALID_RANGE',
      });
      expect(book.revision).toBe(mergedRevision);
      expect(await sheet.range('A1:A3').getValues()).toEqual([[10], [20], [30]]);
      await book.undo();
      expect(book.sheetData(sheet.id).merges).toEqual([]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
});

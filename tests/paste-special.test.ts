import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';

describe('selective cell copy and paste', () => {
  it('rejects stale source snapshots and read-only writes, including hidden rows in normal copies', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A2').setValues([[1], [2]]);
      await sheet.setRowHidden(1, true);
      await sheet.range('A1:A2').copyTo('B1:B2', { mode: 'values' });
      expect(await sheet.range('B1:B2').getValues()).toEqual([[1], [2]]);
      book.setReadOnly(true);
      await expect(sheet.range('A1:A2').copyTo('C1:C2', { mode: 'values' })).rejects.toMatchObject({
        code: 'READ_ONLY',
      });
      book.setReadOnly(false);
      const read = book.readCells.bind(book);
      const spy = vi.spyOn(book, 'readCells').mockImplementationOnce(async (...args) => {
        const data = await read(...args);
        await sheet.range('C1').setValues([[99]]);
        return data;
      });
      await expect(sheet.range('A1:A2').copyTo('C1:C2', { mode: 'values' })).rejects.toMatchObject({
        code: 'REVISION_CONFLICT',
      });
      spy.mockRestore();
      expect(await sheet.range('C1:C2').getValues()).toEqual([[99], [null]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
  it('copies evaluated values, including errors and blanks, without formulas or source formats', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A2').setFormulas([['1+2'], ['1/0']]);
      await sheet.range('A1').setStyle({ bold: true });
      await sheet.range('C1:C3').setValues([[99], [99], [99]]);
      await sheet.range('C1').setStyle({ italic: true });
      await sheet.range('C1').setNote('target');
      await sheet.range('A1:A3').copyTo('C1:C3', { mode: 'values' });
      expect(await sheet.range('C1:C3').getValues()).toEqual([[3], [{ error: '#DIV/0!' }], [null]]);
      expect(await sheet.range('C1:C3').getFormulas()).toEqual([[null], [null], [null]]);
      expect(book.getStyle((await book.getCell(sheet.id, 0, 2))?.styleId)).toEqual({
        italic: true,
      });
      expect(await sheet.range('C1').getNotes()).toEqual([['target']]);
      await book.undo();
      expect(await sheet.range('C1:C3').getValues()).toEqual([[99], [99], [99]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('copies formula inputs with mixed references using a stable snapshot for overlapping ranges', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[7]]);
      await sheet.range('B1').setFormulas([['A1+$A1+A$1+$A$1']]);
      await sheet.range('C1').setStyle({ bold: true });
      await sheet.range('A1:B1').copyTo('B1:C1', { mode: 'formulas' });
      expect(await sheet.range('B1').getValues()).toEqual([[7]]);
      expect(await sheet.range('C1').getFormulas()).toEqual([['(((B1+$A1)+B$1)+$A$1)']]);
      expect(await sheet.range('C1').getValues()).toEqual([[28]]);
      expect(book.getStyle((await book.getCell(sheet.id, 0, 2))?.styleId)).toEqual({ bold: true });
      await book.undo();
      expect(await sheet.range('B1').getFormulas()).toEqual([['A1+$A1+A$1+$A$1']]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('replaces formats, clears absent formats, preserves target content and persists style tables', async () => {
    let book = await createWorkbook();
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1').setStyle({ bold: true, numberFormat: '0.00' });
      await sheet.range('C1:D1').setStyle({ italic: true, background: '#ff0000' });
      await sheet.range('C1:D1').setValues([[3, 4]]);
      await sheet.range('C1').setNote('keep');
      await sheet.range('A1:B1').copyTo('C1:D1', { mode: 'formats' });
      expect(await sheet.range('C1:D1').getValues()).toEqual([[3, 4]]);
      expect(book.getStyle((await book.getCell(sheet.id, 0, 2))?.styleId)).toEqual({
        bold: true,
        numberFormat: '0.00',
      });
      expect(book.getStyle((await book.getCell(sheet.id, 0, 3))?.styleId)).toEqual({});
      expect(await sheet.range('C1').getNotes()).toEqual([['keep']]);
      await book.undo();
      expect(book.getStyle((await book.getCell(sheet.id, 0, 2))?.styleId)).toEqual({
        italic: true,
        background: '#ff0000',
      });
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      sheet = book.getSheets()[0];
      expect(await sheet.range('C1').getDisplayValues()).toEqual([['3.00']]);
      expect(book.getStyle((await book.getCell(sheet.id, 0, 2))?.styleId)).toEqual({
        bold: true,
        numberFormat: '0.00',
      });
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('rolls back new styles and content on invalid targets and rejects malformed payloads', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('B1:C1').merge();
      const before = await book.toJSON(),
        revision = book.revision;
      await expect(
        book.execute({
          type: 'core.cells.paste',
          payload: {
            sheetId: sheet.id,
            cells: [
              {
                row: 0,
                column: 0,
                style: { color: '#123456' },
                input: { type: 'number', value: 7 },
              },
              { row: 0, column: 1, input: { type: 'number', value: 8 } },
            ],
          },
        }),
      ).rejects.toMatchObject({ code: 'INVALID_RANGE' });
      expect(book.revision).toBe(revision);
      expect(await book.toJSON()).toEqual(before);
      await expect(
        book.execute({
          type: 'core.cells.paste',
          payload: { sheetId: sheet.id, cells: [{ row: 0, column: 0 }] },
        }),
      ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
      await expect(sheet.range('B1:C1').copyTo('D1:E1', { mode: 'formats' })).rejects.toMatchObject(
        { code: 'INVALID_RANGE' },
      );
      await expect(sheet.range('A1').copyTo('D1:E1', { mode: 'values' })).rejects.toMatchObject({
        code: 'INVALID_RANGE',
      });
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
});

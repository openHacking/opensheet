import { expect, it } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';

it('pastes annotations atomically with explicit clear and history', async () => {
  const book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0],
      range = sheet.range('A1');
    await range.setValues([[7]]);
    await range.setNote('old');
    await range.setLink({ target: 'https://example.com/old' });
    await book.execute({
      type: 'core.cells.paste',
      payload: { sheetId: sheet.id, cells: [{ row: 0, column: 0, note: 'copied', link: null }] },
    });
    expect(await range.getValues()).toEqual([[7]]);
    expect(await range.getNotes()).toEqual([['copied']]);
    expect(await range.getLinks()).toEqual([[null]]);
    await book.undo();
    expect(await range.getNotes()).toEqual([['old']]);
    expect(await range.getLinks()).toEqual([[{ target: 'https://example.com/old' }]]);
    await book.redo();
    expect(await range.getNotes()).toEqual([['copied']]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('rejects unsafe links and oversized notes before any pasted input or annotation changes', async () => {
  const book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[7, 8]]);
    const before = await book.toJSON(),
      revision = book.revision;
    for (const invalid of [
      { link: { target: 'javascript:alert(1)' } },
      { note: 'x'.repeat(32769) },
    ]) {
      await expect(
        book.execute({
          type: 'core.cells.paste',
          payload: {
            sheetId: sheet.id,
            cells: [
              { row: 0, column: 0, input: { type: 'number', value: 99 }, note: 'must roll back' },
              { row: 0, column: 1, ...invalid },
            ],
          },
        }),
      ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
      expect(book.revision).toBe(revision);
      expect(await book.toJSON()).toEqual(before);
    }
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('preserves styles and annotations across mixed paste transactions and reopen', async () => {
  let book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setStyle({ italic: true });
    await book.transaction({}, async (tx) => {
      const id = tx.getSheets()[0].id;
      await tx.execute({
        type: 'core.cells.paste',
        payload: {
          sheetId: id,
          cells: [{ row: 0, column: 0, input: { type: 'number', value: 42 }, note: 'note' }],
        },
      });
      await tx.execute({
        type: 'core.cells.paste',
        payload: { sheetId: id, cells: [{ row: 0, column: 1, style: { color: '#123456' } }] },
      });
    });
    const verify = async () => {
      const current = book.getSheets()[0];
      expect(await current.range('A1').getValues()).toEqual([[42]]);
      expect(await current.range('A1').getNotes()).toEqual([['note']]);
      expect(book.getStyle((await book.getCell(current.id, 0, 0))?.styleId)).toEqual({
        italic: true,
      });
      expect(book.getStyle((await book.getCell(current.id, 0, 1))?.styleId)).toEqual({
        color: '#123456',
      });
    };
    await verify();
    await book.undo();
    expect(await sheet.range('A1').getValues()).toEqual([[null]]);
    expect(await sheet.range('A1').getNotes()).toEqual([[null]]);
    expect(book.getStyle((await book.getCell(sheet.id, 0, 1))?.styleId)).toEqual({});
    await book.redo();
    await verify();
    const id = book.id;
    await book.close();
    book = await openWorkbook(id);
    await verify();
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

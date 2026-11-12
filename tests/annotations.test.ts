import { describe, expect, it } from 'vitest';
import { createWorkbook, openWorkbook, isSafeLinkTarget } from '@opensheetjs/core';
import { createSnapshot, validateSnapshot } from '../packages/core/src/model.js';

describe('cell notes and links', () => {
  it('preserves annotations through copy, structure, sort, history, native files and reopen', async () => {
    let book = await createWorkbook();
    const target = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B3').setValues([
        [2, 'two'],
        [1, 'one'],
        [3, 'three'],
      ]);
      const link = { target: 'https://example.com/a?q=1#part', tooltip: 'Source' };
      await sheet.range('B1').setNote('备注\nsecond line <script>literal</script>');
      await sheet.range('B1').setLink(link);
      await sheet.range('B2').setNote('first');
      await sheet.range('B1').copyTo('D1');
      expect(await sheet.range('D1').getLinks()).toEqual([[link]]);
      await sheet.insertRows(0);
      expect(await sheet.range('D2').getNotes()).toEqual([
        ['备注\nsecond line <script>literal</script>'],
      ]);
      await sheet.sort('A2:B4', 0);
      expect(await sheet.range('B2:B4').getNotes()).toEqual([
        ['first'],
        ['备注\nsecond line <script>literal</script>'],
        [null],
      ]);
      expect(await sheet.range('B3').getLinks()).toEqual([[link]]);
      await book.undo();
      expect(await sheet.range('B2').getLinks()).toEqual([[link]]);
      await book.redo();
      await target.importJSON(await book.toJSON());
      expect(await target.getSheets()[0].range('B3').getLinks()).toEqual([[link]]);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      expect(await book.getSheets()[0].range('B3').getNotes()).toEqual([
        ['备注\nsecond line <script>literal</script>'],
      ]);
      expect(await book.getSheets()[0].range('B3').getLinks()).toEqual([[link]]);
    } finally {
      await book.close();
      await target.close();
    }
  });

  it('updates note and link independently, and clears values separately from metadata', async () => {
    const book = await createWorkbook();
    try {
      const range = book.getSheets()[0].range('A1');
      await range.setValues([[7]]);
      await range.setNote('note');
      await range.setLink({ target: 'mailto:hello@example.com' });
      await range.setNote(null);
      expect(await range.getLinks()).toEqual([[{ target: 'mailto:hello@example.com' }]]);
      await range.setNote('restored');
      await range.clear();
      expect(await range.getValues()).toEqual([[null]]);
      expect(await range.getNotes()).toEqual([['restored']]);
      await range.clear({ all: true });
      expect(await range.getNotes()).toEqual([[null]]);
      expect(await range.getLinks()).toEqual([[null]]);
      await book.undo();
      expect(await range.getNotes()).toEqual([['restored']]);
    } finally {
      await book.close();
    }
  });

  it.each([
    'javascript:alert(1)',
    'data:text/html,test',
    '//example.com',
    'file:///tmp/x',
    'https://',
    'https:\\example.com',
    ' https://example.com',
    'https://example.com/\npath',
    'mailto:',
  ])('rejects unsafe or malformed target %j atomically', async (target) => {
    expect(isSafeLinkTarget(target)).toBe(false);
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0],
        revision = book.revision;
      await expect(
        book.execute({
          type: 'core.cells.annotation',
          payload: {
            sheetId: sheet.id,
            range: sheet.range('A1').bounds,
            note: 'must not persist',
            link: { target },
          },
        }),
      ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
      expect(book.revision).toBe(revision);
      expect(await sheet.range('A1').getNotes()).toEqual([[null]]);
    } finally {
      await book.close();
    }
  });

  it('rejects covered merged cells, oversized notes and empty annotation payloads', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('B1:C1').merge();
      const revision = book.revision;
      await expect(sheet.range('A1:C1').setNote('partial')).rejects.toMatchObject({
        code: 'INVALID_RANGE',
      });
      expect(await sheet.range('A1').getNotes()).toEqual([[null]]);
      await expect(sheet.range('A1').setNote('x'.repeat(32769))).rejects.toThrow();
      await expect(
        book.execute({
          type: 'core.cells.annotation',
          payload: { sheetId: sheet.id, range: sheet.range('A1').bounds },
        }),
      ).rejects.toThrow();
      expect(book.revision).toBe(revision);
      await sheet.range('B1').setNote('anchor');
      expect(await sheet.range('B1').getNotes()).toEqual([['anchor']]);
    } finally {
      await book.close();
    }
  });

  it('validates native links and accepts explicit approved schemes', () => {
    for (const target of [
      'https://example.com',
      'HTTP://example.com/a',
      'mailto:a@example.com?subject=Hello%20world',
    ])
      expect(isSafeLinkTarget(target)).toBe(true);
    const snapshot = createSnapshot();
    snapshot.sheets[0].cells['0:0'] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'blank' },
      link: { target: 'javascript:alert(1)' },
    };
    expect(() => validateSnapshot(snapshot)).toThrow();
  });
});

it('retains an explicit empty note on a blank cell through history and reopening', async () => {
  let book = await createWorkbook();
  try {
    let sheet = book.getSheets()[0];
    await sheet.range('B1').setNote('');
    expect(await sheet.range('B1').getNotes()).toEqual([['']]);
    await book.undo();
    expect(await sheet.range('B1').getNotes()).toEqual([[null]]);
    await book.redo();
    expect(await sheet.range('B1').getNotes()).toEqual([['']]);
    const id = book.id;
    await book.close();
    book = await openWorkbook(id);
    sheet = book.getSheets()[0];
    expect(await sheet.range('B1').getNotes()).toEqual([['']]);
    await sheet.range('B1').setNote(null);
    expect(await sheet.range('B1').getNotes()).toEqual([[null]]);
    await book.undo();
    expect(await sheet.range('B1').getNotes()).toEqual([['']]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

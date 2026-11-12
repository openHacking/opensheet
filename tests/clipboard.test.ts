import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import {
  copyText,
  cutText,
  parseClipboardText,
  pasteText,
} from '../packages/renderer/src/clipboard.js';
import { copyClipboard } from '../packages/renderer/src/copy-clipboard.js';
import { encodeClipboardText } from '../packages/renderer/src/clipboard-text.js';

describe('clipboard text and cut safety', () => {
  it('returns the captured target even if the caller mutates its selection during the write', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      const selected = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 };
      const execute = book.execute.bind(book);
      let release!: () => void;
      const gate = new Promise<void>((resolve) => {
        release = resolve;
      });
      vi.spyOn(book, 'execute').mockImplementation(async (command) => {
        const result = await execute(command);
        await gate;
        return result;
      });
      const pending = pasteText(book, selected, '5');
      selected.startRow = 10;
      selected.startColumn = 10;
      release();
      expect(await pending).toEqual({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 2,
        startColumn: 0,
        endColumn: 2,
      });
      expect(await sheet.range('A1:B2').getValues()).toEqual([
        [5, 5],
        [5, 5],
      ]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('bounds expansion of long clipboard values before dispatching a tiled paste', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      const execute = vi.spyOn(book, 'execute');
      const revision = book.revision;
      await expect(
        pasteText(
          book,
          { sheetId: sheet.id, startRow: 0, endRow: 1024, startColumn: 0, endColumn: 1 },
          'x'.repeat(32768),
        ),
      ).rejects.toMatchObject({
        code: 'LIMIT_EXCEEDED',
        message: 'Expanded clipboard content is too large',
      });
      expect(execute).not.toHaveBeenCalled();
      expect(book.revision).toBe(revision);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('parses quoted tabs, CR/LF, doubled quotes, trailing blanks and ragged rows', () => {
    expect(parseClipboardText('"a\tb"\t"one\r\ntwo"\r\n"say ""hi"""\t\r\n')).toEqual([
      ['a\tb', 'one\r\ntwo'],
      ['say "hi"', ''],
    ]);
    expect(parseClipboardText('')).toEqual([['']]);
    expect(parseClipboardText('a\tb\n\nlast')).toEqual([['a', 'b'], [''], ['last']]);
    expect(parseClipboardText('a\n""')).toEqual([['a'], ['']]);
    expect(() => parseClipboardText('"unterminated')).toThrow('Unclosed quote');
  });

  it('rejects content after a closed quoted field without altering literal interior quotes', () => {
    for (const input of ['"a"b', '""x', '"a" "b"', 'ok\t"b"c', '"a""b"c'])
      expect(() => parseClipboardText(input)).toThrow('Unexpected text after closing quote');
    expect(parseClipboardText('say "hello"\t"a""b"')).toEqual([['say "hello"', 'a"b']]);
  });

  it('round-trips rectangular clipboard text with quotes, delimiters and trailing blanks', () => {
    const values = ['', 'plain', '"quoted"', 'a\tb', 'a\nb', 'a\r\nb', 'x"y', '😀', ' '];
    let seed = 71237;
    const next = (limit: number) => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return seed % limit;
    };
    for (let sample = 0; sample < 500; sample++) {
      const rows = 1 + next(4),
        columns = 1 + next(4);
      const cells = Array.from({ length: rows }, () =>
        Array.from({ length: columns }, () => values[next(values.length)]),
      );
      expect(parseClipboardText(encodeClipboardText(cells))).toEqual(cells);
    }
  });

  it('pastes a final quoted blank row and restores it on undo', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A2').setValues([['old'], ['preserve on undo']]);
      const result = await pasteText(
        book,
        { sheetId: sheet.id, ...sheet.range('A1').bounds },
        'new\n""',
      );
      expect(result.endRow).toBe(2);
      expect(await sheet.range('A1:A2').getValues()).toEqual([['new'], [null]]);
      await book.undo();
      expect(await sheet.range('A1:A2').getValues()).toEqual([['old'], ['preserve on undo']]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('bounds input dimensions and cell length before rectangular matrix allocation', () => {
    for (const text of [
      'x'.repeat(2 * 1024 * 1024 + 1),
      'x'.repeat(32769),
      '\t'.repeat(1000),
      '\n'.repeat(100001),
      'x\n'.repeat(10000) + '\t'.repeat(20),
    ]) {
      expect(() => parseClipboardText(text)).toThrowError(
        expect.objectContaining({ code: 'LIMIT_EXCEEDED' }),
      );
    }
  });

  it('copies control characters losslessly and pastes one atomic, undoable matrix', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      const values = [
        ['one\rtwo', 'three\nfour'],
        ['tab\there', 'say "hi"'],
      ];
      await sheet.range('A1:B2').setValues(values);
      const text = await copyText(sheet.range('A1:B2'));
      expect(parseClipboardText(text)).toEqual(values);
      const selected = { sheetId: sheet.id, ...sheet.range('D1').bounds };
      const result = await pasteText(book, selected, text);
      expect(result).toMatchObject({ startRow: 0, endRow: 2, startColumn: 3, endColumn: 5 });
      expect(await sheet.range('D1:E2').getValues()).toEqual(values);
      await book.undo();
      expect(await sheet.range('D1:E2').getValues()).toEqual([
        [null, null],
        [null, null],
      ]);
      await sheet.range('D2:E2').merge();
      const revision = book.revision;
      await expect(pasteText(book, selected, '1\t2\n3\t4')).rejects.toMatchObject({
        code: 'INVALID_RANGE',
      });
      expect(book.revision).toBe(revision);
      expect(await sheet.range('D1:E2').getValues()).toEqual([
        [null, null],
        [null, null],
      ]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('copies trailing blank single-column rows so paste clears them and undo restores them', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A3').setValues([['first'], [null], [null]]);
      await sheet.range('C1:C3').setValues([['old'], ['keep on undo'], ['also restore']]);
      const text = await copyText(sheet.range('A1:A3'));
      expect(parseClipboardText(text)).toEqual([['first'], [''], ['']]);
      expect((await copyClipboard(sheet.range('A1:A3'))).text).toBe(text);
      expect(await copyText(sheet.range('A3'))).toBe('');
      const selection = await pasteText(
        book,
        { sheetId: sheet.id, ...sheet.range('C1').bounds },
        text,
      );
      expect(selection.endRow).toBe(3);
      expect(await sheet.range('C1:C3').getValues()).toEqual([['first'], [null], [null]]);
      await book.undo();
      expect(await sheet.range('C1:C3').getValues()).toEqual([
        ['old'],
        ['keep on undo'],
        ['also restore'],
      ]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('never clears source when clipboard access fails, and clears only after write resolves', async () => {
    const book = await createWorkbook();
    try {
      const range = book.getSheets()[0].range('A1');
      await range.setValues([[42]]);
      const revision = book.revision;
      await expect(
        cutText(book, range, async () => {
          throw new Error('clipboard denied');
        }),
      ).rejects.toThrow('clipboard denied');
      expect(book.revision).toBe(revision);
      expect(await range.getValues()).toEqual([[42]]);
      await cutText(book, range, async (text) => {
        expect(text).toBe('42');
        expect(await range.getValues()).toEqual([[42]]);
      });
      expect(book.revision).toBe(revision + 1);
      expect(await range.getValues()).toEqual([[null]]);
      await book.undo();
      expect(await range.getValues()).toEqual([[42]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('rejects clipboard reads that begin before a structural revision changes', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[42]]);
      const range = sheet.range('A1');
      let resume!: () => void;
      const gate = new Promise<void>((resolve) => (resume = resolve));
      const original = book.readRange.bind(book);
      const delayed = vi.spyOn(book, 'readRange').mockImplementation(async (...args) => {
        await gate;
        return original(...args);
      });
      const write = vi.fn(async () => {});
      const cutting = cutText(book, range, write);
      await sheet.insertRows(0);
      resume();
      await expect(cutting).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
      delayed.mockRestore();
      expect(write).not.toHaveBeenCalled();
      expect(await sheet.range('A2').getValues()).toEqual([[42]]);

      let resumeStructured!: () => void;
      const structuredGate = new Promise<void>((resolve) => (resumeStructured = resolve));
      const readCells = book.readCells.bind(book);
      const delayedCells = vi.spyOn(book, 'readCells').mockImplementation(async (...args) => {
        await structuredGate;
        return readCells(...args);
      });
      const copying = copyClipboard(sheet.range('A2'));
      await sheet.insertRows(0);
      resumeStructured();
      await expect(copying).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
      delayedCells.mockRestore();
      expect(await sheet.range('A3').getValues()).toEqual([[42]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('refuses to copy or cut incomplete calculations while preserving literal error-like text', async () => {
    let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 2 }] });
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[1]]);
      await sheet
        .range('A2:A130')
        .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 1}+1`]));
      await sheet.range('B1:B2').setValues([['#LIMIT!'], ['ordinary']]);
      await sheet.range('B2').setFormulas([['1/0']]);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      sheet = book.getSheets()[0];
      const source = sheet.range('A130');
      const revision = book.revision;
      const write = vi.fn(async () => {});
      await expect(cutText(book, source, write)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
      expect(write).not.toHaveBeenCalled();
      expect(book.revision).toBe(revision);
      expect(await source.getFormulas()).toEqual([['A129+1']]);
      expect(await copyText(sheet.range('B1:B2'))).toBe('#LIMIT!\n#DIV/0!');
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('keeps newer edits when the workbook changes during the clipboard write', async () => {
    const book = await createWorkbook();
    try {
      const range = book.getSheets()[0].range('A1');
      await range.setValues([[42]]);
      await expect(
        cutText(book, range, async (text) => {
          expect(text).toBe('42');
          await range.setValues([[99]]);
        }),
      ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
      expect(await range.getValues()).toEqual([[99]]);
      await book.undo();
      expect(await range.getValues()).toEqual([[42]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('does not clear the source if the workbook becomes read-only during the clipboard write', async () => {
    const book = await createWorkbook();
    try {
      const range = book.getSheets()[0].range('A1');
      await range.setValues([[42]]);
      const revision = book.revision;
      await expect(cutText(book, range, async () => book.setReadOnly(true))).rejects.toMatchObject({
        code: 'READ_ONLY',
      });
      expect(book.revision).toBe(revision);
      expect(await range.getValues()).toEqual([[42]]);
      book.setReadOnly(false);
      await book.undo();
      expect(await range.getValues()).toEqual([[null]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('does not clear the source when the editor guard changes during the clipboard write', async () => {
    const book = await createWorkbook();
    try {
      const range = book.getSheets()[0].range('A1');
      await range.setValues([[42]]);
      const revision = book.revision;
      let editable = true;
      await expect(
        cutText(
          book,
          range,
          async () => {
            editable = false;
          },
          () => editable,
        ),
      ).rejects.toMatchObject({ code: 'READ_ONLY' });
      expect(book.isReadOnly).toBe(false);
      expect(book.revision).toBe(revision);
      expect(await range.getValues()).toEqual([[42]]);
      await book.undo();
      expect(await range.getValues()).toEqual([[null]]);
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
});

it('rejects plain-text paste into a merged anchor without changing data or history', async () => {
  const book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([['original']]);
    await sheet.range('A1:B1').merge();
    const revision = book.revision;
    const selected = { sheetId: sheet.id, ...sheet.range('A1').bounds };
    await expect(pasteText(book, selected, 'replacement')).rejects.toMatchObject({
      code: 'INVALID_RANGE',
    });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1:B1').getValues()).toEqual([['original', null]]);
    await book.undo();
    expect(book.sheetData(sheet.id).merges).toEqual([]);
    expect(await sheet.range('A1').getValues()).toEqual([['original']]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

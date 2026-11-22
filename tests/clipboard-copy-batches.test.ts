import { expect, it } from 'vitest';
import { createWorkbook, type Range } from '@opensheetjs/core';
import { copyText, cutText } from '../packages/renderer/src/clipboard.js';
import { copyClipboard } from '../packages/renderer/src/copy-clipboard.js';
import {
  clipboardTextCellLength,
  encodeClipboardText,
} from '../packages/renderer/src/clipboard-text.js';

it('charges exact clipboard encoding including quotes and a trailing blank row', () => {
  for (const rows of [
    [
      ['a"b', 'x\ny'],
      ['', 'z'],
    ],
    [['x'], ['']],
    [['"\t"']],
  ]) {
    let total = 0;
    rows.forEach((row, r) =>
      row.forEach((value, c) => {
        total +=
          clipboardTextCellLength(
            value,
            rows.length > 1 && r === rows.length - 1 && row.length === 1 && value === '',
          ) + (c ? 1 : r ? 1 : 0);
      }),
    );
    expect(total).toBe(encodeClipboardText(rows).length);
  }
});

it('copies wide and tall ranges in bounded row-major batches', async () => {
  const book = await createWorkbook({
    sheets: [
      { name: 'Wide', rows: 4, columns: 300 },
      { name: 'Tall', rows: 300, columns: 2 },
    ],
  });
  try {
    const [wide, tall] = book.getSheets();
    const rows = Array.from({ length: 3 }, (_, r) =>
      Array.from({ length: 300 }, (_, c) => `${r}:${c}${c === 256 ? '\t"' : ''}`),
    );
    await wide.range('A1:KN3').setValues(rows);
    const tallRows = Array.from({ length: 300 }, (_, i) => [String(i), i === 299 ? '' : 'x']);
    await tall.range('A1:B300').setValues(tallRows);
    const originalRange = book.readRange.bind(book);
    const originalCells = book.readCells.bind(book);
    const sizes: number[] = [];
    book.readRange = async (sheet, bounds, values) => {
      sizes.push((bounds.endRow - bounds.startRow) * (bounds.endColumn - bounds.startColumn));
      return originalRange(sheet, bounds, values);
    };
    book.readCells = async (sheet, bounds, options) => {
      sizes.push((bounds.endRow - bounds.startRow) * (bounds.endColumn - bounds.startColumn));
      return originalCells(sheet, bounds, options);
    };
    expect(await copyText(wide.range('A1:KN3'))).toBe(encodeClipboardText(rows));
    expect((await copyClipboard(wide.range('A1:KN3'))).text).toBe(encodeClipboardText(rows));
    expect(await copyText(tall.range('A1:B300'))).toBe(encodeClipboardText(tallRows));
    expect(sizes.length).toBeGreaterThan(3);
    expect(Math.max(...sizes)).toBeLessThanOrEqual(256);
  } finally {
    await book.deleteStorage();
  }
});

it('rejects a changed revision between copy batches before writing or cutting', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 300, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A300').setValues(Array.from({ length: 300 }, (_, i) => [i]));
    const source = sheet.range('A1:A300');
    const original = book.readRange.bind(book);
    let reads = 0;
    book.readRange = async (...args) => {
      if (++reads === 2) await sheet.range('A300').setValues([[999]]);
      return original(...args);
    };
    let writes = 0;
    await expect(
      cutText(book, source, async () => {
        writes++;
      }),
    ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
    expect(writes).toBe(0);
    expect(await sheet.range('A1').getValues()).toEqual([[0]]);
    expect(await sheet.range('A300').getValues()).toEqual([[999]]);
  } finally {
    await book.deleteStorage();
  }
});

it('limits long plain text before clipboard write and keeps the cut source untouched', async () => {
  const value = '"'.repeat(32768);
  let reads = 0;
  let writes = 0;
  let clears = 0;
  const book = {
    revision: 4,
    id: 'fixture',
    sheetData: () => ({ rowOrder: Array.from({ length: 800 }, (_, i) => i), columnOrder: [0] }),
    readRange: async (_sheet: string, range: { startRow: number; endRow: number }) => {
      reads++;
      return {
        revision: 4,
        calculated: {},
        display: Object.fromEntries(
          Array.from({ length: range.endRow - range.startRow }, (_, i) => [
            `${range.startRow + i}:0`,
            value,
          ]),
        ),
      };
    },
    execute: async () => {
      clears++;
    },
  };
  const range = {
    sheet: { id: 'sheet', workbook: book },
    bounds: { startRow: 0, endRow: 800, startColumn: 0, endColumn: 1 },
  } as unknown as Range;
  await expect(
    cutText(book as any, range, async () => {
      writes++;
    }),
  ).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  expect({ reads, writes, clears }).toEqual({ reads: 2, writes: 0, clears: 0 });
});

it('rejects a changed revision between structured clipboard batches', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 300, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A300').setValues(Array.from({ length: 300 }, (_, i) => [i]));
    const original = book.readCells.bind(book);
    let reads = 0;
    book.readCells = async (...args) => {
      if (++reads === 2) await sheet.range('A300').setValues([[999]]);
      return original(...args);
    };
    await expect(copyClipboard(sheet.range('A1:A300'))).rejects.toMatchObject({
      code: 'REVISION_CONFLICT',
    });
    expect(reads).toBe(2);
    expect(await sheet.range('A1').getValues()).toEqual([[0]]);
    expect(await sheet.range('A300').getValues()).toEqual([[999]]);
  } finally {
    await book.deleteStorage();
  }
});

it('rejects ranges over the existing cell limit before starting a read', async () => {
  let reads = 0;
  const book = {
    revision: 0,
    sheetData: () => ({ rowOrder: Array.from({ length: 16385 }, (_, i) => i), columnOrder: [0] }),
    readRange: async () => {
      reads++;
      throw new Error('Unexpected read');
    },
  };
  const range = {
    sheet: { id: 'sheet', workbook: book },
    bounds: { startRow: 0, endRow: 16385, startColumn: 0, endColumn: 1 },
  } as unknown as Range;
  await expect(copyText(range)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  expect(reads).toBe(0);
});

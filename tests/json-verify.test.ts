import { expect, it, vi } from 'vitest';
import { createWorkbook, verifyJSON } from '@opensheetjs/core';
import { Database } from '../packages/core/src/storage.js';

it('verifies replayable JSON with reordered metadata and no database writes', async () => {
  const book = await createWorkbook({
    sheets: [
      { name: 'First', rows: 130, columns: 2 },
      { name: 'Second', rows: 10, columns: 2 },
    ],
  });
  try {
    const [first, second] = book.getSheets();
    await first.range('A1').setValues([[1]]);
    await first.range('A65').setValues([['text']]);
    await second.range('A1').setValues([[true]]);
    const file = await book.toJSON();
    const reordered = {
      sheets: file.sheets,
      ...Object.fromEntries(Object.entries(file).filter(([key]) => key !== 'sheets')),
    };
    const text = JSON.stringify(reordered);
    const write = vi.spyOn(Database.prototype, 'write');
    try {
      const expected = { workbookId: book.id, sheets: 2, pages: 3, cells: 3 };
      expect(await verifyJSON(text)).toEqual(expected);
      expect(await verifyJSON(new Blob([text]))).toEqual(expected);
      await expect(verifyJSON(new ReadableStream() as unknown as Blob)).rejects.toMatchObject({
        code: 'INVALID_ARGUMENT',
      });
      const controller = new AbortController();
      controller.abort();
      await expect(verifyJSON(text, { signal: controller.signal })).rejects.toMatchObject({
        name: 'AbortError',
      });
      expect(write).not.toHaveBeenCalled();
    } finally {
      write.mockRestore();
    }
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('rejects duplicate pages, bad axis membership, unknown nested fields and truncated JSON', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 10, columns: 2 }] });
  try {
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[91]]);
    const file = await book.toJSON();
    for (const mutate of [
      (value: any) => value.sheets[0].blocks.push(structuredClone(value.sheets[0].blocks[0])),
      (value: any) => {
        value.sheets[0].blocks[0].data.numbers.positions[0] = 640;
      },
      (value: any) => {
        value.sheets[0].freeze.unknown = true;
      },
      (value: any) => {
        value.schemaVersion = 4;
      },
    ]) {
      const invalid = structuredClone(file);
      mutate(invalid);
      await expect(verifyJSON(JSON.stringify(invalid))).rejects.toThrow();
    }
    await expect(verifyJSON(JSON.stringify(file).slice(0, -1))).rejects.toThrow();
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

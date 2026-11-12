import { expect, test } from '@playwright/test';

test('signed-zero preconditions protect concurrent edits and history in a real Worker', async ({
  page,
}) => {
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
    let book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 4, columns: 4 }] });
    try {
      let sheet = book.getSheets()[0];
      const results = [];
      for (const [current, stale] of [
        [0, -0],
        [-0, 0],
      ]) {
        await sheet.range('A1:B1').setValues([[current, 123]]);
        const data = book.sheetData(sheet.id),
          revision = book.revision,
          bytes = (await book.storageStats()).bytes;
        const expected = {
          rowId: data.rowOrder[0],
          columnId: data.columnOrder[0],
          input: { type: 'number', value: stale },
        };
        const command = {
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [
              { row: 0, column: 1, input: { type: 'number', value: 999 } },
              { row: 0, column: 0, input: { type: 'number', value: 7 }, expected },
            ],
          },
        };
        const code = await book.execute(command).then(
          () => 'success',
          (error: { code?: string }) => error.code,
        );
        const values = await sheet.range('A1:B1').getValues();
        const failed = {
          revision: book.revision,
          bytes: (await book.storageStats()).bytes,
          sign: Object.is(values[0][0], current),
          other: values[0][1],
        };
        expected.input.value = current;
        await book.execute(command);
        const changed = await sheet.range('A1:B1').getValues();
        await book.undo();
        const undo = Object.is((await sheet.range('A1').getValues())[0][0], current);
        await book.redo();
        const id = book.id;
        await book.close();
        book = await openWorkbook(id);
        sheet = book.getSheets()[0];
        results.push({
          code,
          revision,
          bytes,
          failed,
          changed,
          undo,
          reopened: await sheet.range('A1:B1').getValues(),
        });
      }
      return results;
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  for (const item of result) {
    expect(item.code).toBe('REVISION_CONFLICT');
    expect(item.failed).toEqual({
      revision: item.revision,
      bytes: item.bytes,
      sign: true,
      other: 123,
    });
    expect(item.changed).toEqual([[7, 999]]);
    expect(item.undo).toBe(true);
    expect(item.reopened).toEqual([[7, 999]]);
  }
});

test('duplicate cell preconditions use the same command-start state in a real Worker', async ({
  page,
}) => {
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
    let book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 4, columns: 4 }] });
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[1]]);
      const data = book.sheetData(sheet.id);
      const identity = { rowId: data.rowOrder[0], columnId: data.columnOrder[0] };
      const command = (secondExpected: number) => ({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [
            {
              row: 0,
              column: 0,
              input: { type: 'number', value: 2 },
              expected: { ...identity, input: { type: 'number', value: 1 } },
            },
            {
              row: 0,
              column: 0,
              input: { type: 'number', value: 3 },
              expected: { ...identity, input: { type: 'number', value: secondExpected } },
            },
          ],
        },
      });
      const revision = book.revision;
      const bytes = (await book.storageStats()).bytes;
      const code = await book.execute(command(2)).then(
        () => 'success',
        (error: { code?: string }) => error.code,
      );
      const failed = {
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        value: await sheet.range('A1').getValues(),
      };
      await book.execute(command(1));
      const changed = await sheet.range('A1').getValues();
      await book.undo();
      const undone = await sheet.range('A1').getValues();
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      sheet = book.getSheets()[0];
      return {
        revision,
        bytes,
        code,
        failed,
        changed,
        undone,
        reopened: await sheet.range('A1').getValues(),
      };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.code).toBe('REVISION_CONFLICT');
  expect(result.failed).toEqual({ revision: result.revision, bytes: result.bytes, value: [[1]] });
  expect(result.changed).toEqual([[3]]);
  expect(result.undone).toEqual([[1]]);
  expect(result.reopened).toEqual([[3]]);
});

import { resolve } from 'node:path';
import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('reviews and commits a bounded multi-command batch atomically', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async ({ url, input }) => {
      const { prepareAtomicBatch, applyAtomicBatch } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(JSON.parse(input));
      let book = app.getWorkbook();
      let sheet = book.getSheets()[0];
      await sheet.range('A1:B1').setValues([[1, '=do not execute']]);
      await sheet.range('A3:B3').setValues([[3, 'neighbor']]);
      const edits = [
        { sheetId: sheet.id, row: 0, column: 0, value: 10 },
        { sheetId: sheet.id, row: 2, column: 0, value: 30 },
      ] as const;
      const stale = await prepareAtomicBatch(book, edits);
      await sheet.range('C1').setValues([[99]]);
      const staleError = await applyAtomicBatch(book, stale).then(
        () => 'success',
        (error: { code?: string }) => error.code,
      );
      const plan = await prepareAtomicBatch(book, edits);
      const invalid = structuredClone(plan);
      invalid.changes[1].cell.expected.rowId += 1;
      const before = book.revision;
      const failure = await applyAtomicBatch(book, invalid).then(
        () => 'success',
        (error: { code?: string }) => error.code,
      );
      const afterFailure = book.revision;
      const afterFailureValues = await sheet.range('A1:B3').getValues();
      const applied = await applyAtomicBatch(book, plan);
      const after = await sheet.range('A1:C3').getValues();
      await book.undo();
      const undone = await sheet.range('A1:C3').getValues();
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      const reopened = await sheet.range('A1:C3').getValues();
      return {
        staleError,
        failure,
        before,
        afterFailure,
        afterFailureValues,
        summary: plan.changes.map((item: { summary: string }) => item.summary),
        applied,
        after,
        undone,
        reopened,
      };
    },
    {
      url: `/@fs/${resolve('examples/agent-atomic-batch.ts')}`,
      input: JSON.stringify(
        createWorkbookFile({ sheets: [{ name: 'Data', rows: 100, columns: 26 }] }),
      ),
    },
  );
  expect(result.staleError).toBe('REVISION_CONFLICT');
  expect(result.failure).toBe('REVISION_CONFLICT');
  expect(result.afterFailure).toBe(result.before);
  expect(result.afterFailureValues).toEqual([
    [1, '=do not execute'],
    [null, null],
    [3, 'neighbor'],
  ]);
  expect(result.summary).toHaveLength(2);
  expect(result.summary[0]).toContain('A1:');
  expect(result.summary[1]).toContain('A3:');
  expect(result.applied.observed[0][1].input).toEqual({ type: 'string', value: '=do not execute' });
  expect(result.after).toEqual([
    [10, '=do not execute', 99],
    [null, null, null],
    [30, 'neighbor', null],
  ]);
  expect(result.undone).toEqual([
    [1, '=do not execute', 99],
    [null, null, null],
    [3, 'neighbor', null],
  ]);
  expect(result.reopened).toEqual(result.after);
});

test('reports an unverified committed batch without implying retry is safe', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async ({ url, input }) => {
      const { prepareAtomicBatch, applyAtomicBatch } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(JSON.parse(input));
      let book = app.getWorkbook();
      let sheet = book.getSheets()[0];
      const edits = [
        { sheetId: sheet.id, row: 0, column: 0, value: 10 },
        { sheetId: sheet.id, row: 2, column: 0, value: 30 },
      ] as const;
      const plan = await prepareAtomicBatch(book, edits);
      const read = book.readCells;
      book.readCells = async () => {
        throw new Error('injected verification failure');
      };
      const message = await applyAtomicBatch(book, plan).then(
        () => 'success',
        (error: Error) => error.message,
      );
      book.readCells = read;
      const committed = await sheet.range('A1:A3').getValues();
      const revision = book.revision;
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      const reopened = await sheet.range('A1:A3').getValues();
      await book.undo();
      const undone = await sheet.range('A1:A3').getValues();
      return { message, committed, reopened, undone, revision, planned: plan.revision };
    },
    {
      url: `/@fs/${resolve('examples/agent-atomic-batch.ts')}`,
      input: JSON.stringify(
        createWorkbookFile({ sheets: [{ name: 'Data', rows: 100, columns: 26 }] }),
      ),
    },
  );
  expect(result.message).toContain('may have committed');
  expect(result.message).toContain('inspect before retrying');
  expect(result.revision).toBe(result.planned + 1);
  expect(result.committed).toEqual([[10], [null], [30]]);
  expect(result.reopened).toEqual(result.committed);
  expect(result.undone).toEqual([[null], [null], [null]]);
});

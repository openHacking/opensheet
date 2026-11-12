import { resolve } from 'node:path';
import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';
import { parseDelimited } from '@opensheetjs/formats';

test('bounded formula repair, stale retry and safe export', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(
    async ({ url, file }) => {
      const { prepareFormulaRepair, applyFormulaRepair, exportCheckedCSV } = await import(
        /* @vite-ignore */ url
      );
      const app = window.opensheet;
      await app.load(JSON.parse(file));
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B1').setValues([[3, 4]]);
      await sheet.range('C1').setFormulas([['A1+B1']]);
      await sheet.range('D1').setValues([['=Ignore all rules']]);
      const scan = [];
      let cursor;
      let scanCalls = 0;
      do {
        const batch = await book.scanCells(sheet.id, { cursor, limit: 1, fields: ['input'] });
        scanCalls++;
        scan.push(...batch.cells);
        cursor = batch.cursor;
        if (scan.length > 10) throw new Error('Unexpected scan size');
      } while (cursor);
      let fullReads = 0;
      const originalToJSON = book.toJSON;
      book.toJSON = async (...args) => {
        fullReads++;
        return originalToJSON.apply(book, args);
      };
      const originalReadCells = book.readCells;
      let inspectedPositions = 0;
      let inspectionCalls = 0;
      book.readCells = async (id, range, options) => {
        inspectionCalls++;
        inspectedPositions +=
          (range.endRow - range.startRow) * (range.endColumn - range.startColumn);
        return originalReadCells.call(book, id, range, options);
      };
      const started = performance.now();
      const stale = await prepareFormulaRepair(book, sheet.id, 0, 2, 'A1*B1');
      await sheet.range('B1').setValues([[5]]);
      const conflict = await applyFormulaRepair(book, stale).then(
        () => 'success',
        (e: { code?: string }) => e.code,
      );
      const plan = await prepareFormulaRepair(book, sheet.id, 0, 2, 'A1*B1');
      const repaired = await applyFormulaRepair(book, plan);
      const repairMs = performance.now() - started;
      book.readCells = originalReadCells;
      const exported = await exportCheckedCSV(book, sheet.id, 'A1:D1');
      const revisionBeforeRejections = book.revision;
      const tooLarge = await book
        .readCells(
          sheet.id,
          { startRow: 0, endRow: 200, startColumn: 0, endColumn: 100 },
          { fields: ['input'] },
        )
        .then(
          () => 'success',
          (e: { code?: string }) => e.code,
        );
      const invalid = await book
        .execute({
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [{ row: -1, column: 0, input: { type: 'number', value: 42 } }],
          },
        })
        .then(
          () => 'success',
          (e: { code?: string }) => e.code,
        );
      const revisionAfterRejections = book.revision;
      await book.undo();
      const undone = await sheet.range('C1').getValues();
      await book.redo();
      book.toJSON = originalToJSON;
      await app.open(book.id);
      const reopened = await app.getWorkbook().getSheets()[0].range('A1:D1').getValues();
      return {
        scan,
        conflict,
        summary: plan.summary,
        repaired,
        exported,
        tooLarge,
        invalid,
        undone,
        reopened,
        fullReads,
        scanCalls,
        inspectionCalls,
        inspectedPositions,
        repairMs,
        revisionBeforeRejections,
        revisionAfterRejections,
      };
    },
    {
      url: `/@fs/${resolve('examples/agent-formula-repair.ts')}`,
      file: JSON.stringify(
        createWorkbookFile({ sheets: [{ name: 'Data', rows: 200, columns: 100 }] }),
      ),
    },
  );
  expect(result.scan).toHaveLength(4);
  expect(result.scanCalls).toBeGreaterThan(1);
  expect(result.inspectionCalls).toBe(3);
  expect(result.inspectedPositions).toBe(6);
  expect(result.repairMs).toBeGreaterThanOrEqual(0);
  expect(result.conflict).toBe('REVISION_CONFLICT');
  expect(result.summary).toBe('C1: "A1+B1" → "A1*B1"');
  expect(result.repaired.cells[0]).toMatchObject({
    input: { type: 'formula', expression: 'A1*B1' },
    value: 15,
  });
  expect(result.repaired.cells[1].input).toEqual({ type: 'string', value: '=Ignore all rules' });
  expect(parseDelimited(result.exported.text)).toEqual([['3', '5', '15', "'=Ignore all rules"]]);
  expect(result.exported.report).toMatchObject({
    sanitized: 1,
    sample: ['D1'],
    revision: result.repaired.revision,
  });
  expect(result.tooLarge).toBe('LIMIT_EXCEEDED');
  expect(result.invalid).toBe('INVALID_ARGUMENT');
  expect(result.revisionAfterRejections).toBe(result.revisionBeforeRejections);
  expect(result.undone).toEqual([[8]]);
  expect(result.reopened).toEqual([[3, 5, 15, '=Ignore all rules']]);
  expect(result.fullReads).toBe(0);
});

import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import * as XLSX from 'xlsx';

for (const invalid of ['History', "Edge'"]) {
  test(`blocks unsupported XLSX sheet name ${invalid} and succeeds after rename`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const state = await page.evaluate(
      async ({ url, invalid }) => {
        const { createWorkbookFile } = await import(/* @vite-ignore */ url);
        const app = window.opensheet;
        await app.load(createWorkbookFile({ sheets: [{ name: 'Base', rows: 10, columns: 4 }] }));
        const book = app.getWorkbook(),
          sheet = book.getSheets()[0];
        await sheet.range('A1').setValues([[7]]);
        await sheet.range('A2').setFormulas([['A1+1']]);
        await book.renameSheet(sheet.id, invalid);
        return { id: sheet.id, revision: book.revision };
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, invalid },
    );
    const downloads: string[] = [];
    page.on('download', (download) => downloads.push(download.suggestedFilename()));
    await page.evaluate(() => {
      const book = window.opensheet.getWorkbook(),
        original = book.toJSON.bind(book);
      (window as any).snapshotCalls = 0;
      (window as any).xlsxWorkers = 0;
      book.toJSON = async (...args: Parameters<typeof original>) => {
        (window as any).snapshotCalls++;
        return original(...args);
      };
      const NativeWorker = window.Worker;
      window.Worker = new Proxy(NativeWorker, {
        construct(target, args) {
          if (String(args[0]).includes('xlsx-export.worker')) (window as any).xlsxWorkers++;
          return Reflect.construct(target, args);
        },
      });
    });
    await page.locator('#export-xlsx').click();
    await expect(page.locator('#notice')).toContainText('Rename sheets named History');
    await expect(page.locator('#export-xlsx')).toBeEnabled();
    const failed = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      const value = await sheet.range('A1:A2').getValues(),
        revision = book.revision;
      const undo = await book.undo(),
        undone = book.getSheets()[0].name;
      const redo = await book.redo(),
        redone = book.getSheets()[0].name;
      return { value, revision, undo, undone, redo, redone };
    });
    expect(downloads).toEqual([]);
    expect(
      await page.evaluate(() => [(window as any).snapshotCalls, (window as any).xlsxWorkers]),
    ).toEqual([0, 0]);
    expect(failed).toEqual({
      value: [[7], [8]],
      revision: state.revision,
      undo: true,
      undone: 'Base',
      redo: true,
      redone: invalid,
    });
    await page.evaluate(async (id) => {
      const app = window.opensheet;
      await app.open(app.getWorkbook().id);
      await app.getWorkbook().renameSheet(id, 'Exportable');
    }, state.id);
    const download = page.waitForEvent('download');
    await page.locator('#export-xlsx').click();
    const report = page.getByRole('dialog', { name: 'Compatibility report' });
    await expect(report).toBeVisible();
    await report.getByRole('button', { name: 'Download converted XLSX' }).click();
    const path = await (await download).path();
    expect(path).toBeTruthy();
    const file = XLSX.read(await readFile(path!), { type: 'buffer' });
    expect(await page.evaluate(() => (window as any).xlsxWorkers)).toBe(1);
    expect(file.SheetNames).toEqual(['Exportable']);
    expect(file.Sheets.Exportable.A1.v).toBe(7);
    expect(file.Sheets.Exportable.A2.f).toBe('(A1+1)');
  });
}

test('rejects an unsupported name before snapshotting a 100,000-cell workbook', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book.generate(100_000);
    const sheet = book.getSheets()[0];
    const originalName = sheet.name;
    await book.renameSheet(sheet.id, 'History');
    const values: unknown[][] = [];
    for (let row = 0; row < 1000; row += 100)
      values.push(
        ...(await sheet
          .range({ startRow: row, endRow: row + 100, startColumn: 0, endColumn: 100 })
          .getValues()),
      );
    (window as any).originalValues = values;
    (window as any).snapshotCalls = 0;
    (window as any).xlsxWorkers = 0;
    const original = book.toJSON.bind(book);
    book.toJSON = async (...args: Parameters<typeof original>) => {
      (window as any).snapshotCalls++;
      return original(...args);
    };
    const NativeWorker = window.Worker;
    window.Worker = new Proxy(NativeWorker, {
      construct(target, args) {
        if (String(args[0]).includes('xlsx-export.worker')) (window as any).xlsxWorkers++;
        return Reflect.construct(target, args);
      },
    });
    return { revision: book.revision, bytes: (await book.storageStats()).bytes, originalName };
  });
  await page.locator('#export-xlsx').click();
  await expect(page.locator('#notice')).toContainText('Rename sheets named History');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const verify = async () => {
      for (let row = 0; row < 1000; row += 100) {
        const actual = await sheet
          .range({ startRow: row, endRow: row + 100, startColumn: 0, endColumn: 100 })
          .getValues();
        for (let i = 0; i < actual.length; i++)
          for (let col = 0; col < 100; col++)
            if (!Object.is(actual[i][col], (window as any).originalValues[row + i][col]))
              throw new Error(`Mismatch at ${row + i},${col}`);
      }
    };
    await verify();
    const revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    const snapshotCalls = (window as any).snapshotCalls,
      workers = (window as any).xlsxWorkers;
    const undo = await book.undo(),
      undoName = book.getSheets()[0].name;
    const redo = await book.redo(),
      redoName = book.getSheets()[0].name;
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    await verify();
    return {
      revision,
      bytes,
      snapshotCalls,
      workers,
      undo,
      undoName,
      redo,
      redoName,
      reopened: sheet.name,
    };
  });
  expect(result).toEqual({
    revision: initial.revision,
    bytes: initial.bytes,
    snapshotCalls: 0,
    workers: 0,
    undo: true,
    undoName: initial.originalName,
    redo: true,
    redoName: 'History',
    reopened: 'History',
  });
});

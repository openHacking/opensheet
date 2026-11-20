import { workbookTool } from './workbook-tools.js';
import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import * as XLSX from 'xlsx';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 100, columns: 26 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 3, endRow: 4, startColumn: 2, endColumn: 3 });
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
});

test('freezes before the selection and preserves columns when toggling the toolbar row', async ({
  page,
}) => {
  const freeze = () =>
    page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      return book.sheetData(book.getSheets()[0].id).freeze;
    });
  await workbookTool(page, 'freeze-selection');
  await expect.poll(freeze).toEqual({ rows: 3, columns: 2 });
  const geometry = await page.evaluate(async () => {
    const app = window.opensheet;
    const grid = app.getGrid() as any;
    await grid.ready();
    const layout = grid.layoutState;
    const freeze = app.getWorkbook().sheetData(app.getWorkbook().getSheets()[0].id).freeze;
    const before = {
      fixed: layout.position(0, 0, freeze, grid.scroller),
      moving: layout.position(5, 4, freeze, grid.scroller),
    };
    grid.scroller.scrollTop = 120;
    grid.scroller.scrollLeft = 256;
    await grid.ready();
    return {
      before,
      after: {
        fixed: layout.position(0, 0, freeze, grid.scroller),
        moving: layout.position(5, 4, freeze, grid.scroller),
      },
      scrollTop: grid.scroller.scrollTop,
      scrollLeft: grid.scroller.scrollLeft,
    };
  });
  expect(geometry.scrollTop).toBeGreaterThan(0);
  expect(geometry.scrollLeft).toBeGreaterThan(0);
  expect(geometry.after.fixed).toEqual(geometry.before.fixed);
  expect(geometry.after.moving.x).toBe(geometry.before.moving.x - geometry.scrollLeft);
  expect(geometry.after.moving.y).toBe(geometry.before.moving.y - geometry.scrollTop);
  await workbookTool(page, 'freeze-first-row');
  await expect.poll(freeze).toEqual({ rows: 1, columns: 2 });
  await workbookTool(page, 'unfreeze');
  await expect.poll(freeze).toEqual({ rows: 0, columns: 0 });
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.undo();
    const undone = book.sheetData(book.getSheets()[0].id).freeze;
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { undone, reopened: book.sheetData(book.getSheets()[0].id).freeze };
  });
  expect(result).toEqual({ undone: { rows: 1, columns: 2 }, reopened: { rows: 0, columns: 0 } });
});

test('freezes first axes independently, exports XLSX, and rejects read-only changes', async ({
  page,
}) => {
  const freeze = () =>
    page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      return book.sheetData(book.getSheets()[0].id).freeze;
    });
  await workbookTool(page, 'freeze-first-row');
  await expect.poll(freeze).toEqual({ rows: 1, columns: 0 });
  await workbookTool(page, 'freeze-first-column');
  await expect.poll(freeze).toEqual({ rows: 1, columns: 1 });
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await workbookTool(page, 'freeze-first-row');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  const download = page.waitForEvent('download');
  await page.locator('#export-xlsx').click();
  const exported = await download;
  const path = await exported.path();
  expect(path).toBeTruthy();
  const parsed = XLSX.read(await readFile(path!), { type: 'buffer', bookFiles: true });
  expect(fromSheetJS(parsed).file.sheets[0].freeze).toEqual({ rows: 1, columns: 1 });
  await page.getByLabel('Read only', { exact: true }).check();
  await workbookTool(page, 'unfreeze', true);
  expect(await freeze()).toEqual({ rows: 1, columns: 1 });
});

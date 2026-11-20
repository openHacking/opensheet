import { expect, test, type Page } from '@playwright/test';
import { toolbarTool } from './workbook-tools.js';

async function ready(page: Page) {
  await page.goto('/#features');
  await expect(page.getByRole('heading', { name: 'All Features' })).toBeVisible();
  await page.evaluate(() => window.opensheet.getGrid()!.ready());
}
const grid = (page: Page) => page.getByRole('grid', { name: 'Spreadsheet', exact: true });
const item = (page: Page, name: string) => page.getByRole('menuitem', { name, exact: true });
const selection = (page: Page) => page.evaluate(() => window.opensheet.selection.get());

async function clickGrid(
  page: Page,
  x: number,
  y: number,
  button: 'left' | 'right' = 'left',
  shift = false,
) {
  await grid(page).click({ button, position: { x, y }, modifiers: shift ? ['Shift'] : undefined });
}
async function values(page: Page, name: string, address: string) {
  return page.evaluate(
    async ({ name, address }) => {
      try {
        return await window.opensheet
          .getWorkbook()
          .getSheetByName(name)!
          .range(address)
          .getValues();
      } catch (error) {
        if ((error as { code?: string }).code === 'REVISION_CONFLICT') return null;
        throw error;
      }
    },
    { name, address },
  );
}

test('cell, row and column context menus preserve the right target and provide undo', async ({
  page,
}) => {
  await ready(page);
  await expect(page.getByLabel('Sheet actions', { exact: true })).toHaveCount(0);
  await page.getByLabel('Cell address', { exact: true }).fill('A2:B3');
  await page.getByLabel('Cell address', { exact: true }).press('Enter');
  await clickGrid(page, 196, 105, 'right');
  await expect(page.getByRole('menu', { name: 'Cell actions' })).toBeVisible();
  expect(await selection(page)).toMatchObject({
    startRow: 1,
    endRow: 3,
    startColumn: 0,
    endColumn: 2,
  });
  await item(page, 'Note & link…').click();
  await expect(page.getByRole('dialog')).toHaveAccessibleName('Note and link — B3');
  await page.keyboard.press('Escape');

  await clickGrid(page, 15, 72);
  await clickGrid(page, 15, 104, 'left', true);
  await clickGrid(page, 15, 104, 'right');
  await expect(page.getByRole('menu', { name: 'Row actions' })).toBeVisible();
  expect(await selection(page)).toMatchObject({
    startRow: 1,
    endRow: 3,
    startColumn: 0,
    endColumn: 12,
  });
  await item(page, 'Insert 2 rows below').click();
  await expect.poll(() => values(page, 'Orders', 'A4:A6')).toEqual([[null], [null], ['Desk pad']]);
  await page.evaluate(() => window.opensheet.getWorkbook().undo());
  await expect.poll(() => values(page, 'Orders', 'A4')).toEqual([['Desk pad']]);

  await clickGrid(page, 195, 12);
  await clickGrid(page, 322, 12, 'left', true);
  await clickGrid(page, 322, 12, 'right');
  await expect(page.getByRole('menu', { name: 'Column actions' })).toBeVisible();
  expect(await selection(page)).toMatchObject({
    startRow: 0,
    endRow: 100,
    startColumn: 1,
    endColumn: 3,
  });
  await item(page, 'Delete 2 columns').click();
  await expect.poll(() => values(page, 'Orders', 'A1:B1')).toEqual([['Product', 'Unit price']]);
  await page.evaluate(() => window.opensheet.getWorkbook().undo());
  await expect
    .poll(() => values(page, 'Orders', 'A1:C1'))
    .toEqual([['Product', 'Region', 'Units']]);
});

test('column sorting proposes a range and preserves the header, and Filter reflects the active sheet', async ({
  page,
}) => {
  await ready(page);
  await clickGrid(page, 325, 12, 'right');
  await item(page, 'Sort descending by this column').click();
  await expect(page.getByLabel('Sort range')).toHaveValue('A1:G7');
  await expect(page.getByLabel('First selected row contains headers')).toBeChecked();
  await expect(page.getByLabel('Level 1 column')).toHaveValue('2');
  await expect(page.getByLabel('Level 1 order')).toHaveValue('desc');
  await expect.poll(() => values(page, 'Orders', 'C2')).toEqual([[12]]);
  await page.getByRole('button', { name: 'Apply sort' }).click();
  await expect.poll(() => values(page, 'Orders', 'C1:C2')).toEqual([['Units'], [15]]);

  await clickGrid(page, 60, 12, 'right');
  await item(page, 'Filter this column…').click();
  await page.getByLabel('Condition 1 operator').selectOption('eq');
  await page.getByLabel('Condition 1 value', { exact: true }).fill('Notebook');
  await page.getByRole('button', { name: 'Apply filter' }).click();
  await expect(page.getByRole('dialog')).toBeHidden();
  await expect(await toolbarTool(page, 'Filter (active)')).toBeVisible();
  await (await toolbarTool(page, 'Filter (active)')).click();
  await item(page, 'Clear filter').click();
  await expect(await toolbarTool(page, 'Filter')).toBeVisible();
});

test('keyboard and narrow-width menu entry, plus read-only protection', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await ready(page);
  await grid(page).press('Shift+F10');
  await expect(item(page, 'Copy')).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await expect(item(page, 'Paste special')).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(item(page, 'Paste all')).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(item(page, 'Paste special')).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(grid(page)).toBeFocused();

  await page.getByLabel('Read only', { exact: true }).check();
  await (await toolbarTool(page, 'Row actions')).click();
  await expect(page.getByRole('menu', { name: 'Row actions' })).toBeVisible();
  await expect(item(page, 'Insert 1 row above')).toBeDisabled();
  await page.keyboard.press('Escape');
  await (await toolbarTool(page, 'Column actions')).click();
  await expect(item(page, 'Delete 1 column')).toBeDisabled();
});

for (const [label, expectedValue, expectedFormula, expectedBold] of [
  ['Paste all', 100, '(B2*10)', true],
  ['Values only', 10, null, false],
  ['Formula inputs only', 100, '(B2*10)', false],
  ['Formats only', 20, null, true],
] as const) {
  test(`Paste special applies ${label.toLowerCase()} directly from structured ClipboardItem`, async ({
    page,
  }) => {
    await ready(page);
    await page.evaluate(async () => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      const practice = book.getSheetByName('Practice')!;
      await practice.range('B2').setStyle({ bold: true });
      await practice.range('C2').setValues([[20]]);
      app.selection.set({ sheetId: practice.id, ...practice.range('B2').bounds });
      const data = await app.getGrid()!.copyClipboard();
      Object.defineProperty(navigator, 'clipboard', {
        configurable: true,
        value: {
          read: async () => [
            {
              types: ['text/plain', 'text/html'],
              getType: async (type: string) =>
                new Blob([type === 'text/plain' ? data.text : data.html], { type }),
            },
          ],
        },
      });
      app.selection.set({ sheetId: practice.id, ...practice.range('C2').bounds });
    });
    await grid(page).press('Shift+F10');
    await item(page, 'Paste special').click();
    await item(page, label).click();
    await expect(page.locator('.opensheet')).not.toHaveAttribute('aria-busy', 'true');
    await expect(page.getByRole('dialog')).toBeHidden();
    const actual = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheetByName('Practice')!;
      const cell = await book.getCell(sheet.id, 1, 2);
      const range = sheet.range('C2');
      return {
        value: (await range.getValues())[0][0],
        formula: (await range.getFormulas())[0][0],
        bold: !!book.getStyle(cell?.styleId).bold,
      };
    });
    expect(actual).toEqual({ value: expectedValue, formula: expectedFormula, bold: expectedBold });
  });
}

test('denied clipboard access keeps the selected mode, and delayed reads cannot paste after moving', async ({
  page,
}) => {
  await ready(page);
  await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { read: () => Promise.reject(new DOMException('Denied', 'NotAllowedError')) },
    });
  });
  await grid(page).press('Shift+F10');
  await item(page, 'Paste special').click();
  await item(page, 'Values only').click();
  await expect(page.getByLabel('Paste mode')).toHaveValue('values');
  await expect(page.getByLabel('Paste clipboard here')).toBeFocused();
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        read: () =>
          new Promise((resolve) => {
            (window as any).finishMenuClipboard = () =>
              resolve([
                {
                  types: ['text/plain'],
                  getType: async () => new Blob(['unexpected']),
                },
              ]);
          }),
      },
    });
  });
  await grid(page).press('Shift+F10');
  await item(page, 'Paste').click();
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).finishMenuClipboard))
    .toBe('function');
  await page.evaluate(() => {
    const app = window.opensheet;
    const rates = app.getWorkbook().getSheetByName('Rates')!;
    app.selection.set({ sheetId: rates.id, ...rates.range('B2').bounds });
    (window as any).finishMenuClipboard();
  });
  await expect(page.locator('.os-toast')).toContainText('Selection or workbook changed');
  await expect.poll(() => values(page, 'Orders', 'A1')).toEqual([['Product']]);
  await expect.poll(() => values(page, 'Rates', 'B2')).toEqual([[0.1]]);
});

test('header hit testing follows zoom, frozen panes, scroll and hidden axes', async ({ page }) => {
  await ready(page);
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheetByName('Orders')!;
    await sheet.setRowsHidden(1, 1, true);
    await sheet.setColumnsHidden(1, 1, true);
    const currentGrid = app.getGrid()!;
    currentGrid.setZoom(1.25);
    await currentGrid.ready();
    currentGrid.scroller.scrollTop = 38;
    currentGrid.scroller.scrollLeft = 160;
    await currentGrid.ready();
  });
  await clickGrid(page, 15, 82, 'right');
  expect(await selection(page)).toMatchObject({ startRow: 3, endRow: 4 });
  await page.keyboard.press('Escape');
  await clickGrid(page, 240, 12, 'right');
  expect(await selection(page)).toMatchObject({ startColumn: 3, endColumn: 4 });
  await page.keyboard.press('Escape');
  await clickGrid(page, 80, 12, 'right');
  expect(await selection(page)).toMatchObject({ startColumn: 0, endColumn: 1 });
});

test('sheet tab menus rename their own sheet and confirm deletion', async ({ page }) => {
  await ready(page);
  await page.getByRole('button', { name: 'Options for Rates' }).click();
  page.once('dialog', (dialog) => dialog.accept('Pricing'));
  await item(page, 'Rename sheet…').click();
  await expect(page.getByRole('tab', { name: 'Pricing', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Options for Practice' }).click();
  page.once('dialog', (dialog) => dialog.dismiss());
  await item(page, 'Delete sheet…').click();
  await expect(page.getByRole('tab', { name: 'Practice', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Options for Practice' }).click();
  page.once('dialog', (dialog) => dialog.accept());
  await item(page, 'Delete sheet…').click();
  await expect(page.getByRole('tab', { name: 'Practice', exact: true })).toHaveCount(0);
});

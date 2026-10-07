import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import * as XLSX from 'xlsx';
import type { OpenSheet } from 'opensheet';
declare global {
  interface Window {
    opensheet: OpenSheet;
  }
}
test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await page.waitForFunction(
    () =>
      !!window.opensheet?.selection.get() && !!document.querySelector('#scene-title')?.textContent,
  );
  await expect(page.getByRole('grid', { name: 'Spreadsheet', exact: true })).toBeVisible();
  await expect(page.locator('#source')).toContainText('tabular');
});
test('renders the local workbook, formula totals and source preview', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await expect(page.getByRole('heading', { name: 'Budget Simulator' })).toBeVisible();
  const total = await page.evaluate(
    async () => await window.opensheet.getWorkbook().getSheets()[0].range('F11').getValues(),
  );
  expect(total).toEqual([[13097]]);
  await page.getByRole('tab', { name: 'Code export' }).click();
  await page.getByRole('tab', { name: 'Markdown', exact: true }).click();
  await expect(page.locator('#source')).toContainText('| Item |');
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.screenshot({
    path: `test-results/playground-${test.info().project.name}.png`,
    fullPage: true,
  });
  expect(errors).toEqual([]);
});
test('keeps both scrollbar tracks visible and usable beside the canvas', async ({ page }) => {
  const scroller = page.locator('.os-scroll');
  const bounds = await scroller.evaluate((element) => {
    const rect = element.parentElement!.getBoundingClientRect();
    const canvas = document.querySelector('.os-canvas')!.getBoundingClientRect();
    return {
      left: rect.left,
      top: rect.top,
      width: rect.width,
      height: rect.height,
      clientWidth: element.clientWidth,
      clientHeight: element.clientHeight,
      canvasWidth: canvas.width,
      canvasHeight: canvas.height,
      scrollWidth: element.scrollWidth,
      scrollHeight: element.scrollHeight,
    };
  });
  expect(bounds.width - bounds.clientWidth).toBeGreaterThanOrEqual(12);
  expect(bounds.height - bounds.clientHeight).toBeGreaterThanOrEqual(12);
  expect(bounds.canvasWidth).toBe(bounds.clientWidth);
  expect(bounds.canvasHeight).toBe(bounds.clientHeight);
  expect(bounds.scrollWidth).toBeGreaterThan(bounds.clientWidth);
  expect(bounds.scrollHeight).toBeGreaterThan(bounds.clientHeight);
  const selection = await page.evaluate(() => window.opensheet.selection.get());
  await page.mouse.click(bounds.left + bounds.clientWidth - 20, bounds.top + bounds.height - 6);
  await expect.poll(() => scroller.evaluate((element) => element.scrollLeft)).toBeGreaterThan(0);
  await page.mouse.click(bounds.left + bounds.width - 6, bounds.top + bounds.clientHeight - 20);
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toEqual(selection);
  const horizontal = page.getByRole('scrollbar', { name: 'Horizontal scroll' });
  await horizontal.press('Home');
  await expect.poll(() => scroller.evaluate((element) => element.scrollLeft)).toBe(0);
  const thumb = await horizontal.locator('.os-scrollbar-thumb').boundingBox();
  await page.mouse.move(thumb!.x + thumb!.width / 2, thumb!.y + thumb!.height / 2);
  await page.mouse.down();
  await page.mouse.move(thumb!.x + thumb!.width / 2 + 50, thumb!.y + thumb!.height / 2);
  await page.mouse.up();
  await expect.poll(() => scroller.evaluate((element) => element.scrollLeft)).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toEqual(selection);
  await horizontal.dblclick();
  await expect(page.getByRole('textbox', { name: 'Edit cell', exact: true })).toBeHidden();
});
test('edits with the DOM editor, recalculates, undoes and redoes', async ({ page }) => {
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).fill('D2');
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).press('Enter');
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.press('Enter');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('3');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Enter');
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('F2').getValues(),
    ),
  ).toEqual([[7200]]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('D2').getValues(),
    ),
  ).toEqual([[1]]);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('D2').getValues(),
    ),
  ).toEqual([[3]]);
});
test('uses the formula bar, tab navigation and independent worksheets', async ({ page }) => {
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).fill('H2');
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).press('Enter');
  await page.getByRole('textbox', { name: 'Formula bar', exact: true }).fill('=SUM(D2:D10)');
  await page.getByRole('textbox', { name: 'Formula bar', exact: true }).press('Enter');
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('H2').getValues(),
    ),
  ).toEqual([[15]]);
  await page.getByRole('tab', { name: 'Notes', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Formula bar', exact: true })).toHaveValue(
    'OpenSheet',
  );
  await page.getByRole('button', { name: 'Add sheet', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'Sheet1', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
});
test('blocks edits in read-only mode and preserves keyboard selection', async ({ page }) => {
  await page.getByLabel('Read only', { exact: true }).check();
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.focus();
  await grid.press('x');
  await expect(page.getByRole('textbox', { name: 'Edit cell', exact: true })).toBeHidden();
  await grid.press('ArrowRight');
  await expect(page.getByRole('textbox', { name: 'Cell address', exact: true })).toHaveValue('B1');
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['Item']]);
});
test('pastes a matrix atomically and protects IME composition from Enter', async ({ page }) => {
  await page.evaluate(() => {
    const app = window.opensheet,
      id = app.getWorkbook().getSheets()[0].id;
    app.selection.set({ sheetId: id, startRow: 15, endRow: 16, startColumn: 0, endColumn: 1 });
    app.getGrid()!.pasteText('姓名\t数量\n测试\t2');
  });
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('A16:B17').getValues(),
    ),
  ).toEqual([
    ['姓名', '数量'],
    ['测试', 2],
  ]);
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.focus();
  await grid.press('Enter');
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await editor.dispatchEvent('compositionstart');
  await editor.fill('中文输入');
  await editor.dispatchEvent('keydown', { key: 'Enter', isComposing: true });
  await expect(editor).toBeVisible();
  await editor.dispatchEvent('compositionend');
  await editor.press('Enter');
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('A16').getValues(),
    ),
  ).toEqual([['中文输入']]);
});
test('imports a real XLSX in a Worker and exports a downloadable XLSX', async ({ page }) => {
  const wb = XLSX.utils.book_new(
    XLSX.utils.aoa_to_sheet([
      ['Name', 'Count'],
      ['Test', 7],
    ]),
    'Imported',
  );
  const bytes = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  await page.locator('#file').setInputFiles({
    name: 'sample.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  await expect(page.getByRole('tab', { name: 'Imported', exact: true })).toBeVisible();
  await expect(page.locator('#save-state')).toContainText('Imported locally');
  if (await page.getByRole('dialog').isVisible())
    await page.getByRole('button', { name: 'Close dialog', exact: true }).click();
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues(),
    ),
  ).toEqual([[7]]);
  await page.getByRole('button', { name: 'Export XLSX', exact: false }).click();
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download converted XLSX', exact: true }).click();
  expect((await download).suggestedFilename()).toBe('sample.xlsx');
});
test('imports hostile JSON without replacing the current workbook', async ({ page }) => {
  await page.locator('#file').setInputFiles({
    name: 'bad.json',
    mimeType: 'application/json',
    buffer: Buffer.from('{"schemaVersion":99}'),
  });
  await expect(page.locator('#notice')).toContainText('Invalid');
  await expect(page.getByRole('tab', { name: 'Launch budget', exact: true })).toBeVisible();
});
test('provides an accessible data table and a functional plugin', async ({ page }) => {
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 1,
      endRow: 3,
      startColumn: 3,
      endColumn: 4,
    });
  });
  const sum = page.getByRole('button', { name: 'Σ Sum', exact: true });
  if (!(await sum.isVisible()))
    await page.getByRole('button', { name: 'More tools', exact: true }).click();
  await sum.click();
  await expect(page.locator('.os-toast')).toContainText('sum to 2');
  await page.getByRole('button', { name: 'Table view', exact: true }).click();
  await expect(page.getByRole('cell', { name: 'Brand identity', exact: true })).toBeVisible();
});
test('loads saved JSON and cleans up old plugin UI', async ({ page }) => {
  const snapshot = await page.evaluate(async () => await window.opensheet.getWorkbook().toJSON());
  const downloading = page.waitForEvent('download');
  await page.locator('#save-json').click();
  const downloaded = await downloading;
  const buffer = await readFile((await downloaded.path())!);
  const file = JSON.parse(buffer.toString('utf8'));
  expect(file).toEqual(snapshot);
  expect(file.schemaVersion).toBe(3);
  expect(file.sheets[0]).not.toHaveProperty('cells');
  await page.locator('#file').setInputFiles({
    name: 'saved.json',
    mimeType: 'application/json',
    buffer,
  });
  await expect(page.locator('#save-state')).toContainText('Imported locally');
  await expect(
    page.getByRole('button', { name: 'Σ Sum', exact: true, includeHidden: true }),
  ).toHaveCount(1);
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('F11').getValues(),
    ),
  ).toEqual([[13097]]);
});
test('uses a responsive layout without body overflow', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await expect(page.getByRole('grid', { name: 'Spreadsheet', exact: true })).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.screenshot({
    path: `test-results/mobile-${test.info().project.name}.png`,
    fullPage: true,
  });
});
test('scrolls to distant selections and reads the edited frozen-grid cell', async ({ page }) => {
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).fill('J80');
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).press('Enter');
  expect(await page.locator('.os-scroll').evaluate((e) => e.scrollTop)).toBeGreaterThan(1000);
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.press('Enter');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('distant');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Enter');
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('J80').getValues(),
    ),
  ).toEqual([['distant']]);
});
test('cell editor grows for long wrapped and multiline text without visible scrollbars', async ({
  page,
}) => {
  await page.evaluate(() => {
    const app = window.opensheet;
    const sheetId = app.getWorkbook().getSheets()[0].id;
    app.selection.set({ sheetId, startRow: 15, endRow: 16, startColumn: 0, endColumn: 1 });
  });
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.focus();
  await grid.press('Enter');
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await editor.fill(`${'longword'.repeat(24)}\nsecond line`);
  const dimensions = await editor.evaluate((element) => {
    const editor = element as HTMLTextAreaElement;
    const grid = editor.parentElement!;
    const editorRect = editor.getBoundingClientRect();
    const gridRect = grid.getBoundingClientRect();
    return {
      height: editor.clientHeight,
      scrollWidth: editor.scrollWidth,
      clientWidth: editor.clientWidth,
      overflow: getComputedStyle(editor).overflow,
      scrollbarWidth: getComputedStyle(editor).scrollbarWidth,
      within:
        editorRect.left >= gridRect.left &&
        editorRect.right <= gridRect.right + 1 &&
        editorRect.bottom <= gridRect.bottom + 1,
    };
  });
  expect(dimensions.height).toBeGreaterThan(30);
  expect(dimensions.scrollWidth).toBeLessThanOrEqual(dimensions.clientWidth + 1);
  expect(dimensions.scrollbarWidth).toBe('none');
  expect(dimensions.within).toBe(true);
  await page.evaluate(() => window.opensheet.getGrid()!.setZoom(1.5));
  await page.setViewportSize({ width: 390, height: 844 });
  await expect
    .poll(() =>
      editor.evaluate((element) => {
        const rect = element.getBoundingClientRect(),
          grid = element.parentElement!.getBoundingClientRect();
        return (
          rect.left >= grid.left && rect.right <= grid.right + 1 && rect.bottom <= grid.bottom + 1
        );
      }),
    )
    .toBe(true);
  await editor.fill('line\n'.repeat(100));
  expect(
    await editor.evaluate((element) => (element as HTMLTextAreaElement).scrollHeight),
  ).toBeGreaterThan(
    await editor.evaluate((element) => (element as HTMLTextAreaElement).clientHeight),
  );
  await editor.press('Escape');
});
test('merged cells select their full rectangle with pointer, API, and keyboard', async ({
  page,
}) => {
  const point = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('H2').setValues([['merged']]);
    await sheet.range('H2:J4').merge();
    const scroller = document.querySelector('.os-scroll')!;
    scroller.scrollLeft = 800;
    const grid = document.querySelector('.os-grid')!.getBoundingClientRect();
    const data = app.getWorkbook().sheetData(sheet.id);
    const width = (column: number) => data.columns[data.columnOrder[column]]?.size ?? 128;
    const x =
      grid.left +
      48 +
      Array.from({ length: 8 }, (_, i) => width(i)).reduce((a, b) => a + b, 0) +
      width(8) / 2 -
      scroller.scrollLeft;
    const y = grid.top + 30 + 36 + 30 + 15;
    return { x, y, sheetId: sheet.id };
  });
  await page.mouse.click(point.x, point.y);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 1,
    endRow: 4,
    startColumn: 7,
    endColumn: 10,
  });
  await expect
    .poll(() =>
      page.evaluate(() => {
        const app = window.opensheet,
          sheet = app.getWorkbook().getSheets()[0];
        const data = app.getWorkbook().sheetData(sheet.id);
        const left =
          48 +
          Array.from(
            { length: 7 },
            (_, index) => data.columns[data.columnOrder[index]]?.size ?? 128,
          ).reduce((a, b) => a + b, 0) -
          document.querySelector('.os-scroll')!.scrollLeft +
          1;
        const pixel = document
          .querySelector<HTMLCanvasElement>('.os-canvas')!
          .getContext('2d')!
          .getImageData(Math.round(left), 82, 1, 1).data;
        return pixel[1] > pixel[0] && pixel[1] > pixel[2];
      }),
    )
    .toBe(true);
  await expect(page.getByRole('textbox', { name: 'Cell address', exact: true })).toHaveValue(
    'H2:J4',
  );
  await page.evaluate(
    ({ sheetId }) =>
      window.opensheet.selection.set({
        sheetId,
        startRow: 2,
        endRow: 3,
        startColumn: 8,
        endColumn: 9,
      }),
    point,
  );
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 1,
    endRow: 4,
    startColumn: 7,
    endColumn: 10,
  });
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.focus();
  await grid.press('ArrowRight');
  await expect(page.getByRole('textbox', { name: 'Cell address', exact: true })).toHaveValue('K2');
  await grid.press('ArrowLeft');
  await expect(page.getByRole('textbox', { name: 'Cell address', exact: true })).toHaveValue(
    'H2:J4',
  );
  await grid.press('Enter');
  await expect(page.getByRole('textbox', { name: 'Edit cell', exact: true })).toHaveValue('merged');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Escape');
  await page.mouse.move(point.x, point.y);
  await page.mouse.down();
  await page.mouse.move(point.x - 250, point.y + 60, { steps: 5 });
  await page.mouse.up();
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 1,
    endRow: 5,
    startColumn: 6,
    endColumn: 10,
  });
});
test('frozen row and column respect custom header height while scrolling', async ({ page }) => {
  const colors = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setRowHeight(0, 52);
    await sheet.setFreeze(1, 1);
    await sheet.range('A1').setStyle({ background: '#ff0000' });
    await sheet.range('B2').setStyle({ background: '#0000ff' });
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
    const scroller = document.querySelector('.os-scroll')!;
    scroller.scrollTop = 24;
    scroller.scrollLeft = 24;
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    await app.getGrid()!.ready();
    const canvas = document.querySelector<HTMLCanvasElement>('.os-canvas')!;
    const ctx = canvas.getContext('2d')!;
    const pixel = (x: number, y: number) =>
      Array.from(ctx.getImageData(x, y, 1, 1).data).slice(0, 3);
    return {
      frozen: pixel(80, 75),
      body: pixel(260, 84),
      headerFrozen: pixel(10, 75),
      headerBody: pixel(10, 84),
      columnFrozen: pixel(240, 10),
      columnBody: pixel(260, 10),
    };
  });
  expect(colors.frozen).toEqual([255, 0, 0]);
  expect(colors.body[2]).toBeGreaterThan(colors.body[0]);
  expect(colors.headerFrozen).not.toEqual(colors.headerBody);
  expect(colors.columnFrozen).not.toEqual(colors.columnBody);
  const zoomedHeaders = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setRowHidden(1, true);
    app.getGrid()!.setZoom(1.5);
    app.selection.set({ sheetId: sheet.id, startRow: 2, endRow: 3, startColumn: 0, endColumn: 1 });
    document.querySelector('.os-scroll')!.scrollTop = 0;
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const context = document.querySelector<HTMLCanvasElement>('.os-canvas')!.getContext('2d')!;
    return [
      Array.from(context.getImageData(10, 100, 1, 1).data),
      Array.from(context.getImageData(10, 112, 1, 1).data),
    ];
  });
  expect(zoomedHeaders[0]).not.toEqual(zoomedHeaders[1]);
});
test('action icons render as SVG while controls keep accessible names', async ({ page }) => {
  for (const name of ['Undo', 'Redo', 'Merge', 'Freeze first row']) {
    const button = page.getByRole('button', { name, exact: true });
    if (!(await button.isVisible()))
      await page.getByRole('button', { name: 'More tools', exact: true }).click();
    await expect(button.locator('svg')).toHaveCount(1);
  }
  for (const name of ['Import file', 'Export XLSX', 'Keyboard shortcuts']) {
    await expect(page.getByRole('button', { name, exact: true }).locator('svg')).toHaveCount(1);
  }
});
test('toolbar moves trailing actions into More tools instead of scrolling', async ({ page }) => {
  await page.setViewportSize({ width: 600, height: 800 });
  const toolbar = page.getByRole('toolbar', { name: 'Spreadsheet tools' });
  const more = toolbar.getByRole('button', { name: 'More tools' });
  await expect(more).toBeVisible();
  expect(
    await toolbar.evaluate((element) => ({
      overflowX: getComputedStyle(element).overflowX,
      scrollWidth: element.scrollWidth,
      clientWidth: element.clientWidth,
    })),
  ).toMatchObject({ overflowX: 'visible' });
  expect(await toolbar.evaluate((element) => element.scrollWidth <= element.clientWidth)).toBe(
    true,
  );
  await more.click();
  await expect(more).toHaveAttribute('aria-expanded', 'true');
  await expect(toolbar.getByRole('button', { name: 'Freeze first row' })).toBeVisible();
  const freezeRows = await page.evaluate(() => {
    const app = window.opensheet;
    return app.getWorkbook().sheetData(app.getWorkbook().getSheets()[0].id).freeze.rows;
  });
  await toolbar.getByRole('button', { name: 'Freeze first row' }).click();
  await expect(more).toHaveAttribute('aria-expanded', 'false');
  await expect
    .poll(() =>
      page.evaluate(() => {
        const app = window.opensheet;
        return app.getWorkbook().sheetData(app.getWorkbook().getSheets()[0].id).freeze.rows;
      }),
    )
    .toBe(freezeRows ? 0 : 1);
  await page.setViewportSize({ width: 320, height: 800 });
  await expect(more).toBeVisible();
  expect(await toolbar.evaluate((element) => element.scrollWidth <= element.clientWidth)).toBe(
    true,
  );
  await page.setViewportSize({ width: 2400, height: 800 });
  await page.locator('#spreadsheet').evaluate((element) => {
    (element as HTMLElement).style.width = '2200px';
  });
  await expect(more).toBeHidden();
  await expect(toolbar.getByRole('button', { name: 'Freeze first row' })).toBeVisible();
});
test('merged cells crossing frozen panes clip and scroll each pane independently', async ({
  page,
}) => {
  const colors = await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    for (let column = 0; column < 9; column++) await sheet.setColumnWidth(column, 60);
    await sheet.setRowHeight(0, 52);
    await sheet.range('H1').setStyle({ background: '#ff00ff' });
    await sheet.range('H1:I2').merge();
    await sheet.setFreeze(1, 8);
    const scroller = document.querySelector('.os-scroll')!;
    scroller.scrollLeft = 45;
    scroller.scrollTop = 25;
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const context = document.querySelector<HTMLCanvasElement>('.os-canvas')!.getContext('2d')!;
    const pixel = (x: number, y: number) =>
      Array.from(context.getImageData(x, y, 1, 1).data).slice(0, 3);
    return {
      frozen: pixel(480, 55),
      right: pixel(535, 55),
      lower: pixel(480, 84),
      corner: pixel(535, 84),
      beyondRight: pixel(550, 55),
      beyondBottom: pixel(535, 90),
    };
  });
  for (const pane of ['frozen', 'right', 'lower', 'corner'] as const) {
    expect(colors[pane][0]).toBeGreaterThan(200);
    expect(colors[pane][2]).toBeGreaterThan(200);
  }
  expect(colors.beyondRight).toEqual([255, 255, 255]);
  expect(colors.beyondBottom).toEqual([255, 255, 255]);
});
test('replacing workbooks and disposing twice releases component resources', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const NativeResizeObserver = window.ResizeObserver;
    const observers = new Set<ResizeObserver>();
    const globalListeners = new Set<EventListenerOrEventListenerObject>();
    const windowAdd = window.addEventListener.bind(window);
    const documentAdd = document.addEventListener.bind(document);
    const documentRemove = document.removeEventListener.bind(document);
    window.ResizeObserver = class extends NativeResizeObserver {
      constructor(callback: ResizeObserverCallback) {
        super(callback);
        observers.add(this);
      }
      disconnect() {
        observers.delete(this);
        super.disconnect();
      }
    };
    window.addEventListener = ((
      type: string,
      listener: EventListenerOrEventListenerObject,
      options?: AddEventListenerOptions,
    ) => {
      if (type === 'pointermove' || type === 'pointerup') {
        globalListeners.add(listener);
        options?.signal?.addEventListener('abort', () => globalListeners.delete(listener), {
          once: true,
        });
      }
      windowAdd(type, listener, options);
    }) as typeof window.addEventListener;
    document.addEventListener = ((
      type: string,
      listener: EventListenerOrEventListenerObject,
      options?: AddEventListenerOptions,
    ) => {
      if (type === 'pointerdown' || type === 'keydown') globalListeners.add(listener);
      documentAdd(type, listener, options);
    }) as typeof document.addEventListener;
    document.removeEventListener = ((
      type: string,
      listener: EventListenerOrEventListenerObject,
      options?: EventListenerOptions,
    ) => {
      globalListeners.delete(listener);
      documentRemove(type, listener, options);
    }) as typeof document.removeEventListener;
    const container = document.createElement('div');
    container.style.height = '300px';
    document.body.append(container);
    const Editor = window.opensheet.constructor as new (options: {
      container: HTMLElement;
    }) => OpenSheet;
    const editor = new Editor({ container });
    let setups = 0,
      cleanups = 0,
      disposedEvents = 0;
    try {
      await editor.createWorkbook();
      editor.on('lifecycle:disposed', () => {
        disposedEvents++;
        editor.dispose();
      });
      editor.use({
        id: 'test.lifecycle',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: ['ui.toolbar', 'events'],
        setup(ctx) {
          setups++;
          ctx.ui.toolbar.add({ id: 'test.lifecycle.action', label: 'Lifecycle', run() {} });
          ctx.onCommit(() => {});
          ctx.addCleanup(() => {
            cleanups++;
          });
        },
      });
      for (let i = 0; i < 3; i++) await editor.load(await editor.getWorkbook().toJSON());
      editor.notify('Dispose this timer');
      editor.dispose();
      editor.dispose();
      return {
        setups,
        cleanups,
        disposedEvents,
        observers: observers.size,
        listeners: globalListeners.size,
        grids: container.querySelectorAll('.os-grid').length,
      };
    } finally {
      editor.dispose();
      container.remove();
      window.ResizeObserver = NativeResizeObserver;
      window.addEventListener = windowAdd;
      document.addEventListener = documentAdd;
      document.removeEventListener = documentRemove;
    }
  });
  expect(result).toEqual({
    setups: 4,
    cleanups: 4,
    disposedEvents: 1,
    observers: 0,
    listeners: 0,
    grids: 0,
  });
});
test('cancels imports on scene changes and ignores previously queued Worker results', async ({
  page,
}) => {
  await page.evaluate(async () => {
    const workers: Array<{
      terminated: boolean;
      deliver?: (event: MessageEvent) => unknown;
    }> = [];
    Object.assign(window, {
      testWorkers: workers,
      staleSnapshot: await window.opensheet.getWorkbook().toJSON(),
    });
    const NativeWorker = window.Worker;
    window.Worker = class {
      onmessage: ((event: MessageEvent) => unknown) | null = null;
      onerror: (() => unknown) | null = null;
      terminated = false;
      deliver?: (event: MessageEvent) => unknown;
      constructor(url: URL | string, options?: WorkerOptions) {
        if (!String(url).includes('file.worker')) return new NativeWorker(url, options) as any;
        workers.push(this);
      }
      postMessage() {
        this.deliver = this.onmessage ?? undefined;
      }
      terminate() {
        this.terminated = true;
      }
    } as unknown as typeof Worker;
  });
  await page.locator('#file').setInputFiles({
    name: 'pending.csv',
    mimeType: 'text/csv',
    buffer: Buffer.from('a,b\n1,2'),
  });
  await expect(page.getByRole('button', { name: 'Reading…' })).toBeDisabled();
  await expect
    .poll(() => page.evaluate(() => (window as any).testWorkers[0]?.deliver !== undefined))
    .toBe(true);
  await page.getByRole('combobox', { name: 'Choose demo' }).selectOption('sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await page.evaluate(async () => {
    const testWindow = window as any;
    await testWindow.testWorkers[0].deliver({ data: { file: testWindow.staleSnapshot } });
  });
  await expect(page.getByRole('tab', { name: 'Sales Dashboard', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Import file', exact: true })).toBeEnabled();
  expect(await page.evaluate(() => (window as any).testWorkers[0].terminated)).toBe(true);
});
test('sheet switches expand selection using the target sheet merges after removal', async ({
  page,
}) => {
  const selection = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const old = book.getSheets()[0];
    const target = await book.addSheet('Merged target');
    await target.range('A1:B2').merge();
    app.selection.set({ sheetId: target.id, startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });
    const afterSwitch = app.selection.get();
    app.selection.set({ sheetId: old.id, startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });
    await book.removeSheet(old.id);
    app.selection.set({ sheetId: target.id, startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });
    return { targetId: target.id, afterSwitch, afterRemoval: app.selection.get() };
  });
  const expected = {
    sheetId: selection.targetId,
    startRow: 0,
    startColumn: 0,
    endRow: 2,
    endColumn: 2,
  };
  expect(selection.afterSwitch).toEqual(expected);
  expect(selection.afterRemoval).toEqual(expected);
});

test('downloads and reimports binary native files in the real storage worker', async ({ page }) => {
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['native round trip']]);
  });
  const download = page.waitForEvent('download');
  await page.locator('#save-native').click();
  const file = await download;
  const path = await file.path();
  expect(file.suggestedFilename()).toMatch(/\.opensheet$/);
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['changed']]);
  });
  await page.locator('#file').setInputFiles({
    name: file.suggestedFilename(),
    mimeType: 'application/octet-stream',
    buffer: await readFile(path!),
  });
  await expect
    .poll(() =>
      page.evaluate(
        async () => await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    )
    .toEqual([['native round trip']]);
  await expect(page.getByRole('button', { name: 'Import file', exact: true })).toBeEnabled();
});
test('scene changes abort pending native imports before they can reattach the prior workbook', async ({
  page,
}) => {
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    Object.assign(window, { nativeAborted: false });
    book.importJSON = async (_input, options = {}) =>
      await new Promise((_resolve, reject) =>
        options.signal!.addEventListener(
          'abort',
          () => {
            (window as any).nativeAborted = true;
            reject(new Error('cancelled'));
          },
          { once: true },
        ),
      );
  });
  await page.locator('#file').setInputFiles({
    name: 'pending.json',
    mimeType: 'application/json',
    buffer: Buffer.from('{}'),
  });
  await expect(page.getByRole('button', { name: 'Reading…' })).toBeDisabled();
  await page.getByRole('combobox', { name: 'Choose demo' }).selectOption('sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  expect(await page.evaluate(() => (window as any).nativeAborted)).toBe(true);
  await expect(page.getByRole('tab', { name: 'Sales Dashboard', exact: true })).toBeVisible();
});

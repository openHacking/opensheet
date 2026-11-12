import { expect, test } from '@playwright/test';

test('Page Up and Page Down move selection and Shift extends from its anchor', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setFreeze(2, 0);
    await sheet.setRowHidden(5, true);
    app.getGrid()!.setZoom(1.5);
    app.selection.set({ sheetId: sheet.id, startRow: 4, endRow: 5, startColumn: 3, endColumn: 4 });
  });
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  const selection = () => page.evaluate(() => window.opensheet.selection.get()!);
  await page.keyboard.press('PageDown');
  const down = await selection();
  expect(down.startRow).toBeGreaterThan(4);
  expect(down.startRow).not.toBe(5);
  expect(down.endRow).toBe(down.startRow + 1);
  expect(down.startColumn).toBe(3);
  expect(await page.locator('.os-scroll').evaluate((element) => element.scrollTop)).toBeGreaterThan(
    0,
  );
  const visibility = await page.evaluate((row) => {
    const scroller = document.querySelector('.os-scroll')!;
    const top = 30 + (row - (row > 5 ? 1 : 0)) * 45 - scroller.scrollTop;
    return { top, bottom: top + 45, height: scroller.clientHeight };
  }, down.startRow);
  expect(visibility.bottom).toBeLessThanOrEqual(visibility.height);
  await page.keyboard.press('Shift+PageDown');
  const extended = await selection();
  expect(extended.startRow).toBe(down.startRow);
  expect(extended.endRow).toBeGreaterThan(down.endRow);
  expect(extended.startColumn).toBe(3);
  expect(extended.endColumn).toBe(4);
  await page.keyboard.press('Shift+PageUp');
  const shrunk = await selection();
  expect(shrunk.startRow).toBeLessThanOrEqual(down.startRow);
  expect(shrunk.endRow).toBeLessThan(extended.endRow);
  await page.keyboard.press('PageUp');
  const up = await selection();
  expect(up.startRow).toBeLessThan(down.startRow);
  expect(up.endRow).toBe(up.startRow + 1);
  expect(up.startColumn).toBe(3);
});

test('page navigation respects merged cells and a hidden worksheet boundary', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const last = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A10:A12').clear({ all: true });
    await sheet.range('A10:A12').merge();
    await sheet.setRowHeight(9, 75);
    await sheet.setRowHidden(sheet.rowCount - 1, true);
    app.selection.set({ sheetId: sheet.id, startRow: 9, endRow: 10, startColumn: 0, endColumn: 1 });
    return sheet.rowCount - 1;
  });
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await page.keyboard.press('PageDown');
  expect(
    (await page.evaluate(() => window.opensheet.selection.get()!)).startRow,
  ).toBeGreaterThanOrEqual(12);
  await page.evaluate((last) => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: last - 1,
      endRow: last,
      startColumn: 0,
      endColumn: 1,
    });
  }, last);
  await page.keyboard.press('PageDown');
  expect((await page.evaluate(() => window.opensheet.selection.get()!)).startRow).toBe(last - 1);
  await page.keyboard.press('PageUp');
  expect((await page.evaluate(() => window.opensheet.selection.get()!)).startRow).toBeLessThan(
    last - 1,
  );
});

test('page navigation leaves an active cell draft in the editor', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('D5').setValues([['saved']]);
    app.selection.set({ sheetId: sheet.id, startRow: 4, endRow: 5, startColumn: 3, endColumn: 4 });
  });
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await grid.press('Enter');
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await editor.fill('draft');
  await editor.press('PageDown');
  await expect(editor).toHaveValue('draft');
  expect(
    await page.evaluate(async () => ({
      selection: window.opensheet.selection.get(),
      value: await window.opensheet.getWorkbook().getSheets()[0].range('D5').getValues(),
    })),
  ).toMatchObject({
    selection: { startRow: 4, endRow: 5, startColumn: 3, endColumn: 4 },
    value: [['saved']],
  });
});

test('keyboard movement reveals an oversized row below frozen rows', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setFreeze(2, 0);
    await sheet.setRowHeight(10, 1000);
    app.selection.set({ sheetId: sheet.id, startRow: 9, endRow: 10, startColumn: 0, endColumn: 1 });
  });
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await grid.press('ArrowDown');
  const result = await page.evaluate(() => {
    const scroller = document.querySelector('.os-scroll')!;
    return {
      row: window.opensheet.selection.get()!.startRow,
      top: 30 + 10 * 30 - scroller.scrollTop,
      bottom: scroller.clientHeight,
    };
  });
  expect(result.row).toBe(10);
  expect(result.top).toBeGreaterThanOrEqual(90);
  expect(result.top).toBeLessThan(result.bottom);
});

test('keyboard movement reveals a resized column beyond frozen columns', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setFreeze(0, 2);
    for (let column = 0; column < 8; column++) await sheet.setColumnWidth(column, 128);
    await sheet.setColumnWidth(8, 150);
    app.getGrid()!.setZoom(1.5);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 7, endColumn: 8 });
  });
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await grid.press('ArrowRight');
  const result = await page.evaluate(() => {
    const scroller = document.querySelector('.os-scroll')!;
    const left = 48 + 8 * 128 * 1.5 - scroller.scrollLeft;
    return {
      column: window.opensheet.selection.get()!.startColumn,
      left,
      right: left + 150 * 1.5,
      width: scroller.clientWidth,
    };
  });
  expect(result.column).toBe(8);
  expect(result.left).toBeGreaterThanOrEqual(48 + 2 * 128 * 1.5);
  expect(result.right).toBeLessThanOrEqual(result.width);
});

test('programmatic selection reveals its full cell at zoom with frozen panes', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const bounds = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setFreeze(2, 2);
    for (let column = 0; column < 8; column++) await sheet.setColumnWidth(column, 128);
    await sheet.setColumnWidth(8, 150);
    await sheet.setRowHeight(20, 90);
    app.getGrid()!.setZoom(1.5);
    app.selection.set({
      sheetId: sheet.id,
      startRow: 20,
      endRow: 21,
      startColumn: 8,
      endColumn: 9,
    });
    const scroller = document.querySelector('.os-scroll')!;
    const left = 48 + 8 * 128 * 1.5 - scroller.scrollLeft;
    const top = 30 + 20 * 30 * 1.5 - scroller.scrollTop;
    return {
      left,
      right: left + 150 * 1.5,
      top,
      bottom: top + 90 * 1.5,
      width: scroller.clientWidth,
      height: scroller.clientHeight,
    };
  });
  expect(bounds.left).toBeGreaterThanOrEqual(48 + 2 * 128 * 1.5);
  expect(bounds.right).toBeLessThanOrEqual(bounds.width);
  expect(bounds.top).toBeGreaterThanOrEqual(30 + 2 * 30 * 1.5);
  expect(bounds.bottom).toBeLessThanOrEqual(bounds.height);
});

test('selection listener replacement prevents stale keyboard scrolling', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setFreeze(0, 0);
    app.selection.set({
      sheetId: sheet.id,
      startRow: 30,
      endRow: 31,
      startColumn: 0,
      endColumn: 1,
    });
    const unsubscribe = app.selection.onChange(() => {
      unsubscribe();
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
    });
  });
  await page.getByRole('grid', { exact: true }).press('ArrowDown');
  const result = await page.evaluate(() => ({
    selection: window.opensheet.selection.get(),
    scrollTop: document.querySelector('.os-scroll')!.scrollTop,
  }));
  expect(result.selection?.startRow).toBe(0);
  expect(result.scrollTop).toBe(0);
});

test('selection reveals a merged range and aligns an oversized merge', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setFreeze(2, 1);
    for (let column = 0; column < 10; column++) await sheet.setColumnWidth(column, 100);
    await sheet.range('G16:H18').clear({ all: true });
    await sheet.range('G16:H18').merge();
    await sheet.range('A30:A50').clear({ all: true });
    await sheet.range('A30:A50').merge();
    app.getGrid()!.setZoom(1.5);
    app.selection.set({
      sheetId: sheet.id,
      startRow: 15,
      endRow: 16,
      startColumn: 6,
      endColumn: 7,
    });
    const scroller = document.querySelector('.os-scroll')!;
    const left = 48 + 6 * 150 - scroller.scrollLeft;
    const top = 30 + 15 * 45 - scroller.scrollTop;
    const visible = {
      left,
      right: left + 300,
      top,
      bottom: top + 135,
      width: scroller.clientWidth,
      height: scroller.clientHeight,
    };
    app.selection.set({
      sheetId: sheet.id,
      startRow: 29,
      endRow: 30,
      startColumn: 0,
      endColumn: 1,
    });
    return {
      visible,
      oversizedTop: 30 + 29 * 45 - scroller.scrollTop,
      oversizedBottom: scroller.clientHeight,
    };
  });
  expect(result.visible.left).toBeGreaterThanOrEqual(48 + 150);
  expect(result.visible.right).toBeLessThanOrEqual(result.visible.width);
  expect(result.visible.top).toBeGreaterThanOrEqual(30 + 90);
  expect(result.visible.bottom).toBeLessThanOrEqual(result.visible.height);
  expect(result.oversizedTop).toBeGreaterThanOrEqual(30 + 90);
  expect(result.oversizedTop).toBeLessThan(result.oversizedBottom);
});

test('a merged range crossing frozen panes is revealed from selection and keyboard navigation', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('B2:D4').clear({ all: true });
    for (let column = 0; column < 4; column++) await sheet.setColumnWidth(column, 100);
    for (let row = 0; row < 4; row++) await sheet.setRowHeight(row, 30);
    await sheet.range('B2').setStyle({ background: '#ff00ff' });
    await sheet.range('B2:D4').merge();
    await sheet.setFreeze(2, 2);
    app.getGrid()!.setZoom(1.5);
    const scroller = document.querySelector('.os-scroll')!;
    scroller.scrollTop = 500;
    scroller.scrollLeft = 500;
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
  });
  const scroll = () => page.locator('.os-scroll').evaluate((e) => [e.scrollTop, e.scrollLeft]);
  expect(await scroll()).toEqual([0, 0]);
  await page.evaluate(() => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 4, endRow: 5, startColumn: 3, endColumn: 4 });
    const scroller = document.querySelector('.os-scroll')!;
    scroller.scrollTop = 500;
    scroller.scrollLeft = 500;
  });
  await page.getByRole('grid', { exact: true }).press('ArrowUp');
  expect(await scroll()).toEqual([0, 0]);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 1,
    endRow: 4,
    startColumn: 1,
    endColumn: 4,
  });
  const pixels = await page.evaluate(async () => {
    await window.opensheet.getGrid()!.ready();
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
    const ctx = document.querySelector<HTMLCanvasElement>('.os-canvas')!.getContext('2d')!;
    const pixel = (x: number, y: number) =>
      Array.from(
        ctx.getImageData(Math.round(x * devicePixelRatio), Math.round(y * devicePixelRatio), 1, 1)
          .data,
      ).slice(0, 3);
    return {
      inside: [pixel(250, 95), pixel(400, 95), pixel(250, 150), pixel(400, 150)],
      outside: [pixel(670, 150), pixel(400, 220)],
    };
  });
  for (const [red, green, blue] of pixels.inside) {
    expect(red).toBeGreaterThan(200);
    expect(green).toBeLessThan(60);
    expect(blue).toBeGreaterThan(200);
  }
  for (const [red, green, blue] of pixels.outside)
    expect(red > 200 && green < 60 && blue > 200).toBe(false);
});

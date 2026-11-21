import { expect, test } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => window.opensheet.ready());
});

test('formula completion and argument hints never commit before explicit submission', async ({
  page,
}) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=SU');
  await expect(input).toHaveAttribute('aria-expanded', 'true');
  await expect(page.getByRole('option', { name: 'SUM(value, ...)', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await input.press('Enter');
  await expect(input).toHaveValue('=SUM(');
  await expect(input).toBeFocused();
  await expect(page.locator('.os-formula-bar .os-formula-assist [role="status"]')).toContainText(
    'Argument 1',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.pressSequentially('1,2)');
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before + 1);
  const value = await page.evaluate(() => {
    const selection = window.opensheet.selection.get()!;
    return window.opensheet.getWorkbook().getSheetById(selection.sheetId)!.range('A1').getValues();
  });
  expect(value).toEqual([[3]]);
});

test('keyboard selection, quoted text, IME and escape preserve draft semantics', async ({
  page,
}) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=SU');
  await input.press('ArrowDown');
  await input.press('Tab');
  await expect(input).toHaveValue('=SUMIF(');
  await input.fill('=SU(1)');
  await input.evaluate((element: HTMLInputElement) => {
    element.setSelectionRange(3, 3);
    element.dispatchEvent(new Event('input', { bubbles: true }));
  });
  await input.press('Enter');
  await expect(input).toHaveValue('=SUM(1)');
  expect(await input.evaluate((element: HTMLInputElement) => element.selectionStart)).toBe(5);
  await expect(input).toHaveAttribute('aria-expanded', 'false');
  await input.fill('=CONCAT("SU');
  await expect(input).toHaveAttribute('aria-expanded', 'false');
  await input.fill('=SU');
  await input.dispatchEvent('compositionstart');
  await input.dispatchEvent('keydown', { key: 'Enter', isComposing: true, keyCode: 229 });
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.dispatchEvent('compositionend');
  await input.press('Escape');
  await expect(input).toBeFocused();
  await expect(input).toHaveAttribute('aria-expanded', 'false');
  await input.press('Escape');
  await expect(input).not.toBeFocused();
  await expect(input).not.toHaveValue('=SU');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
});

test('formula editing recovers when composition is interrupted by blur', async ({ page }) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=SU');
  await input.dispatchEvent('compositionstart');
  await input.evaluate((element: HTMLInputElement) => element.blur());
  await input.fill('=SU');
  await expect(input).toHaveAttribute('aria-expanded', 'true');
  await input.press('Enter');
  await expect(input).toHaveValue('=SUM(');
  await input.fill('=A1+B2');
  await expect(page.locator('.os-formula-reference-mirror b')).toHaveText(['A1', 'B2']);
  await expect(page.locator('.os-formula-reference-mirror')).toBeVisible();
  await input.press('Escape');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
});

test('dragging a reference keeps the edit target and writes only on explicit commit', async ({
  page,
}) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('D2:D3').setValues([[5], [8]]);
    return { revision: book.revision, references: await sheet.range('D2:D3').getValues() };
  });
  await input.fill('=SUM(');
  const points = await page.evaluate(() => {
    const grid = document.querySelector('.os-grid')!.getBoundingClientRect();
    const book = window.opensheet.getWorkbook();
    const sheet = book.sheetData(book.getSheets()[0].id);
    const x =
      grid.left +
      48 +
      sheet.columnOrder.slice(0, 3).reduce((sum, id) => sum + (sheet.columns[id]?.size ?? 128), 0) +
      8;
    const y = grid.top + 30 + (sheet.rows[sheet.rowOrder[0]]?.size ?? 30) + 8;
    return { x, y, nextY: y + (sheet.rows[sheet.rowOrder[1]]?.size ?? 30) };
  });
  await page.mouse.move(points.x, points.y);
  await page.mouse.down();
  await page.mouse.move(points.x, points.nextY, { steps: 3 });
  await page.mouse.up();
  await expect(input).toBeFocused();
  await expect(input).toHaveValue('=SUM(D2:D3');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before.revision);
  await input.pressSequentially(')');
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before.revision + 1);
  const after = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    return {
      selected: window.opensheet.selection.get(),
      target: await sheet.range('A1').getValues(),
      references: await sheet.range('D2:D3').getValues(),
    };
  });
  expect(after.target).toEqual([[13]]);
  expect(after.references).toEqual(before.references);
  expect(after.selected).toMatchObject({ startRow: 0, startColumn: 0 });
});

test('a changed selection cannot redirect a formula draft and Escape restores its target', async ({
  page,
}) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=1+2');
  await page.evaluate(() => {
    const s = window.opensheet.selection.get()!;
    window.opensheet.selection.set({ ...s, startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
  });
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before + 1);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[3]]);
  await input.fill('=SUM(');
  await page.evaluate(() => {
    const s = window.opensheet.selection.get()!;
    window.opensheet.selection.set({ ...s, startRow: 2, endRow: 3, startColumn: 2, endColumn: 3 });
  });
  await input.press('Escape');
  await expect(input).toHaveValue('=1+2');
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 0,
    startColumn: 0,
  });
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before + 1);
});

test('read-only cancels a draft and replacement preserves it until cancelled', async ({ page }) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const original = await input.inputValue();
  await input.fill('=999');
  await page.evaluate(() => window.opensheet.setMode('read'));
  await expect(input).toHaveValue(original);
  await expect(input).not.toBeFocused();
  await page.evaluate(() => window.opensheet.setMode('edit'));
  await input.fill('=777');
  const result = await page.evaluate(async () => {
    const snapshot = await window.opensheet.getWorkbook().toJSON();
    try {
      await window.opensheet.load(snapshot);
      return 'loaded';
    } catch (error) {
      return (error as { code?: string }).code;
    }
  });
  expect(result).toBe('BUSY');
  await expect(input).toBeFocused();
  await expect(input).toHaveValue('=777');
  await input.press('Escape');
  await page.evaluate(async () => {
    const snapshot = await window.opensheet.getWorkbook().toJSON();
    await window.opensheet.load(snapshot);
  });
  await expect(input).not.toBeFocused();
  await expect(input).toHaveValue(original);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).not.toEqual([[777]]);
});

test('grid clicks inside quoted draft text do not insert references or commit', async ({
  page,
}) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=CONCAT("text');
  const point = await page.evaluate(() => {
    const grid = document.querySelector('.os-grid')!.getBoundingClientRect();
    const book = window.opensheet.getWorkbook();
    const sheet = book.sheetData(book.getSheets()[0].id);
    return {
      x: grid.left + 48 + (sheet.columns[sheet.columnOrder[0]]?.size ?? 128) + 8,
      y: grid.top + 30 + (sheet.rows[sheet.rowOrder[0]]?.size ?? 30) + 8,
    };
  });
  await page.mouse.click(point.x, point.y);
  await expect(input).not.toBeFocused();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 1,
    startColumn: 1,
  });
});

test('cross-sheet references preserve the original target and quote sheet names', async ({
  page,
}) => {
  const setup = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const origin = book.getSheets()[0];
    const other = await book.addSheet("O'Brien Q1");
    await other.range('A1:A2').setValues([[5], [8]]);
    return { originName: origin.name, originId: origin.id, revision: book.revision };
  });
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await input.fill('=SUM(');
  await page.getByRole('tab', { name: "O'Brien Q1", exact: true }).click();
  await expect(input).toBeFocused();
  const grid = await page.locator('.os-grid').boundingBox();
  await page.mouse.move(grid!.x + 48 + 8, grid!.y + 30 + 8);
  await page.mouse.down();
  await page.mouse.move(grid!.x + 48 + 8, grid!.y + 30 + 30 + 8, { steps: 3 });
  await page.mouse.up();
  await expect(input).toHaveValue("=SUM('O''Brien Q1'!A1:A2");
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(setup.revision);
  await input.pressSequentially(')');
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(setup.revision + 1);
  await expect(page.getByRole('tab', { name: setup.originName, exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  const after = await page.evaluate(async ({ originId }) => {
    const book = window.opensheet.getWorkbook();
    return {
      selection: window.opensheet.selection.get(),
      result: await book.getSheetById(originId)!.range('A1').getValues(),
      source: await book
        .getSheets()
        .find((sheet) => sheet.name === "O'Brien Q1")!
        .range('A1:A2')
        .getValues(),
    };
  }, setup);
  expect(after.selection).toMatchObject({ sheetId: setup.originId, startRow: 0, startColumn: 0 });
  expect(after.result).toEqual([[13]]);
  expect(after.source).toEqual([[5], [8]]);
  await input.fill('=SUM(');
  await page.getByRole('tab', { name: "O'Brien Q1", exact: true }).click();
  await input.press('Escape');
  await expect(page.getByRole('tab', { name: setup.originName, exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(
    setup.revision + 1,
  );
});

test('keyboard reference picker changes sheets, extends a range and does not commit on acceptance', async ({
  page,
}) => {
  const setup = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const origin = book.getSheets()[0];
    const other = await book.addSheet("O'Brien Q2");
    await other.range('A1:A2').setValues([[4], [9]]);
    return {
      revision: book.revision,
      originName: origin.name,
      originId: origin.id,
      sheetSteps: book.getSheets().findIndex((sheet) => sheet.id === other.id),
    };
  });
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await input.fill('=SUM(');
  await input.press('Alt+ArrowDown');
  await expect(page.locator('.os-formula-bar .os-formula-assist [role="status"]')).toContainText(
    'Arrows move',
  );
  for (let i = 0; i < setup.sheetSteps; i++) await input.press('PageDown');
  await input.press('Shift+ArrowDown');
  await expect(input).toHaveValue("=SUM('O''Brien Q2'!A1:A2");
  const prevented = await input.evaluate((element: HTMLInputElement) => {
    const event = new KeyboardEvent('keydown', {
      key: 'Enter',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    });
    element.dispatchEvent(event);
    return event.defaultPrevented;
  });
  expect(prevented).toBe(false);
  await expect(page.locator('.os-formula-bar .os-formula-assist [role="status"]')).toContainText(
    'Arrows move',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(setup.revision);
  await input.press('Enter');
  await expect(input).toBeFocused();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(setup.revision);
  await input.pressSequentially(')');
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(setup.revision + 1);
  await expect(page.getByRole('tab', { name: setup.originName, exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  expect(
    await page.evaluate(
      (id) => window.opensheet.getWorkbook().getSheetById(id)!.range('A1').getValues(),
      setup.originId,
    ),
  ).toEqual([[13]]);
  await input.fill('=SUM(');
  await input.press('Alt+ArrowDown');
  for (let i = 0; i < setup.sheetSteps; i++) await input.press('PageDown');
  await input.press('Escape');
  await expect(input).toHaveValue('=SUM(');
  await expect(input).toBeFocused();
  await expect(page.getByRole('tab', { name: setup.originName, exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(
    setup.revision + 1,
  );
  await input.press('Alt+ArrowDown');
  for (let i = 0; i < setup.sheetSteps; i++) await input.press('PageDown');
  for (let i = 0; i < 3; i++) await input.press('ArrowDown');
  await input.press('Shift+ArrowUp');
  await input.press('Shift+ArrowUp');
  await expect(input).toHaveValue("=SUM('O''Brien Q2'!A2:A4");
  await input.press('Escape');
  await expect(input).toHaveValue('=SUM(');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(
    setup.revision + 1,
  );
});

test('formula references draw distinct outlines and clear when editing ends', async ({ page }) => {
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  const colors = () =>
    page.locator('.os-canvas').evaluate((canvas: HTMLCanvasElement) => {
      const pixels = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      let blue = 0,
        orange = 0;
      for (let i = 0; i < pixels.length; i += 4) {
        if (pixels[i] === 37 && pixels[i + 1] === 99 && pixels[i + 2] === 235) blue++;
        if (pixels[i] === 217 && pixels[i + 1] === 119 && pixels[i + 2] === 6) orange++;
      }
      return { blue, orange };
    });
  await input.fill('=SUM(B2,D3)');
  await expect.poll(async () => (await colors()).blue).toBeGreaterThan(10);
  await expect.poll(async () => (await colors()).orange).toBeGreaterThan(10);
  await input.press('Escape');
  await expect.poll(async () => (await colors()).blue).toBe(0);
  await expect.poll(async () => (await colors()).orange).toBe(0);
});

test('reference outlines are clipped around frozen panes and follow the visible sheet', async ({
  page,
}) => {
  await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book.getSheets()[0].setFreeze(1, 1);
  });
  const scroller = page.locator('.os-scroll');
  await scroller.evaluate((element) => {
    element.scrollLeft = 128;
    element.scrollTop = 30;
  });
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await input.fill('=SUM(A1:C3)');
  const bluePixels = () =>
    page.locator('.os-canvas').evaluate((canvas: HTMLCanvasElement) => {
      const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      const scale = window.devicePixelRatio || 1;
      let count = 0,
        outside = 0;
      for (let y = 0; y < canvas.height; y++)
        for (let x = 0; x < canvas.width; x++) {
          const i = (y * canvas.width + x) * 4;
          if (data[i] === 37 && data[i + 1] === 99 && data[i + 2] === 235) {
            count++;
            if (x < 48 * scale || y < 30 * scale) outside++;
          }
        }
      return { count, outside };
    });
  await expect.poll(async () => (await bluePixels()).count).toBeGreaterThan(10);
  expect((await bluePixels()).outside).toBe(0);
  await input.fill("='Notes'!A1");
  await expect.poll(async () => (await bluePixels()).count).toBe(0);
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const other = book.getSheetByName('Notes')!;
    window.opensheet.selection.set({
      sheetId: other.id,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
  });
  await expect(input).toBeFocused();
  await expect.poll(async () => (await bluePixels()).count).toBeGreaterThan(10);
  await input.press('Escape');
  await expect.poll(async () => (await bluePixels()).count).toBe(0);
});

test('a reference outside a sheet shortened during editing does not break the grid', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await input.fill('=SUM(B90)');
  await page.evaluate(async () => {
    const sheet = window.opensheet.getWorkbook().getSheets()[0];
    await sheet.deleteRows(2, sheet.rowCount - 2);
  });
  await expect(input).toBeFocused();
  await expect(input).toHaveValue('=SUM(B90)');
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  expect(errors).toEqual([]);
  await input.press('Escape');
  await expect(page.getByRole('grid', { name: 'Spreadsheet' })).toBeVisible();
});

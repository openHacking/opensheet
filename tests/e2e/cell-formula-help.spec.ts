import { expect, test } from '@playwright/test';

test('cancelling a composing cell edit does not block later grid input', async ({ page }) => {
  const grid = page.getByRole('grid', { name: 'Spreadsheet' });
  const input = page.locator('.os-cell-editor');
  await input.fill('=SU');
  await input.dispatchEvent('compositionstart');
  await page.evaluate(() => window.opensheet.setMode('read'));
  await page.evaluate(() => window.opensheet.setMode('edit'));
  await grid.focus();
  await grid.press('F2');
  await expect(input).toBeVisible();
  await input.fill('=SU');
  await expect(page.locator('.os-grid [role="listbox"]')).toBeVisible();
  await input.press('Escape');
  await input.press('Escape');
  await expect(input).toBeHidden();
});

test('a composition ending after blur commits its final input exactly once', async ({ page }) => {
  const input = page.locator('.os-cell-editor');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('draft');
  await input.dispatchEvent('compositionstart');
  await input.evaluate((element: HTMLTextAreaElement) => {
    element.blur();
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '文' }));
    element.value = 'draft文';
    element.dispatchEvent(new InputEvent('input', { bubbles: true, data: '文' }));
  });
  await expect(input).toBeHidden();
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before + 1);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['draft文']]);
});

test('refocusing an interrupted composition retains the draft for editing', async ({ page }) => {
  const input = page.locator('.os-cell-editor');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('draft');
  await input.dispatchEvent('compositionstart');
  await input.evaluate((element: HTMLTextAreaElement) => {
    element.blur();
    element.focus();
  });
  await input.pressSequentially('!');
  await expect(input).toHaveValue('draft!');
  await input.press('Escape');
  await expect(input).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
});

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
  await page.keyboard.press('F2');
});

test('in-cell references map to outline colors without changing the draft', async ({ page }) => {
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const mirror = page.locator('.os-cell-reference-mirror');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=SUM(A1,\nB2,Missing!A1,C3)');
  await expect(mirror.locator('b')).toHaveText(['A1', 'B2', 'C3']);
  await expect(mirror).toHaveAttribute('aria-hidden', 'true');
  expect(
    await mirror
      .locator('b')
      .evaluateAll((marks) => marks.map((mark) => (mark as HTMLElement).style.textDecorationColor)),
  ).toEqual(['rgb(37, 99, 235)', 'rgb(217, 119, 6)', 'rgb(147, 51, 234)']);
  const geometry = await input.evaluate((element: HTMLTextAreaElement) => {
    const mirror = element.parentElement!.querySelector<HTMLElement>('.os-cell-reference-mirror')!;
    const a = element.getBoundingClientRect(),
      b = mirror.getBoundingClientRect();
    const inputStyle = getComputedStyle(element),
      mirrorStyle = getComputedStyle(mirror);
    return {
      input: [a.x, a.y, a.width, a.height],
      mirror: [b.x, b.y, b.width, b.height],
      fonts: ['font-family', 'font-size', 'font-weight', 'line-height', 'letter-spacing'].map(
        (property) => [
          inputStyle.getPropertyValue(property),
          mirrorStyle.getPropertyValue(property),
        ],
      ),
    };
  });
  expect(geometry.mirror).toEqual(geometry.input);
  for (const [inputFont, mirrorFont] of geometry.fonts) expect(inputFont).toBe(mirrorFont);
  await input.fill(`=SUM(${Array.from({ length: 80 }, () => 'A1').join(',\n')})`);
  const scroll = await input.evaluate((element: HTMLTextAreaElement) => {
    element.scrollTop = element.scrollHeight;
    element.dispatchEvent(new Event('scroll'));
    const mirror = element.parentElement!.querySelector<HTMLElement>('.os-cell-reference-mirror')!;
    return [element.scrollTop, mirror.scrollTop];
  });
  expect(scroll[0]).toBeGreaterThan(0);
  expect(Math.abs(scroll[0] - scroll[1])).toBeLessThanOrEqual(1);
  await input.dispatchEvent('compositionstart');
  await expect(mirror).toBeHidden();
  await input.dispatchEvent('compositionend');
  await expect(mirror).toBeVisible();
  await input.press('Escape');
  await expect(mirror).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
});

test('in-cell reference positions track wrapped text at 200% zoom', async ({ page }) => {
  await page.evaluate(() => {
    document.body.style.zoom = '2';
  });
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await input.fill('=SUM(A1,\nB2,12345678901234567890,C3)');
  const positions = await input.evaluate((element: HTMLTextAreaElement) => {
    const mirror = element.parentElement!.querySelector<HTMLElement>('.os-cell-reference-mirror')!;
    const computed = getComputedStyle(element);
    const plain = document.createElement('div');
    for (const property of [
      'font-family',
      'font-size',
      'font-style',
      'font-weight',
      'line-height',
      'letter-spacing',
      'padding',
      'border',
      'box-sizing',
      'white-space',
      'overflow-wrap',
      'font-kerning',
      'font-variant-ligatures',
      'text-align',
    ])
      plain.style.setProperty(property, computed.getPropertyValue(property));
    Object.assign(plain.style, {
      position: 'absolute',
      left: element.style.left,
      top: element.style.top,
      width: element.style.width,
      height: element.style.height,
      overflow: 'hidden',
      pointerEvents: 'none',
      visibility: 'hidden',
    });
    plain.textContent = element.value;
    element.parentElement!.append(plain);
    plain.scrollTop = element.scrollTop;
    plain.scrollLeft = element.scrollLeft;
    const range = document.createRange();
    const result = [...mirror.querySelectorAll('b')].map((mark) => {
      const start = element.value.indexOf(mark.textContent!);
      range.setStart(plain.firstChild!, start);
      range.setEnd(plain.firstChild!, start + mark.textContent!.length);
      const expected = range.getBoundingClientRect(),
        actual = mark.getBoundingClientRect();
      return [Math.abs(actual.left - expected.left), Math.abs(actual.top - expected.top)];
    });
    plain.remove();
    return result;
  });
  expect(positions).toHaveLength(3);
  for (const [x, y] of positions) {
    expect(x).toBeLessThan(2);
    expect(y).toBeLessThan(2);
  }
});

test('in-cell completion keeps multiline textbox semantics and commits only after confirmation', async ({
  page,
}) => {
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=SU');
  await expect(page.locator('.os-grid [role="listbox"]')).toBeVisible();
  await expect(input).toHaveAttribute('aria-autocomplete', 'list');
  const active = await input.getAttribute('aria-activedescendant');
  const list = await input.getAttribute('aria-controls');
  expect(active).toBeTruthy();
  expect(list).toBeTruthy();
  await expect(page.locator(`#${list} #${active}`)).toHaveText('SUM(value, ...)');
  await input.press('Enter');
  await expect(input).toHaveValue('=SUM(');
  await expect(input).toBeFocused();
  await expect(page.locator('.os-grid [role="status"]')).toContainText('Argument 1');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.pressSequentially('1,2)');
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before + 1);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[3]]);
  await expect(input).toBeHidden();
  await expect(page.locator('.os-grid [role="listbox"]')).toBeHidden();
});

test('mouse completion, IME and multiline input do not accidentally commit', async ({ page }) => {
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=SU');
  await page.locator('.os-grid [role="option"]').first().click();
  await expect(input).toHaveValue('=SUM(');
  await expect(input).toBeFocused();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.fill('=SU');
  await input.dispatchEvent('compositionstart');
  await input.dispatchEvent('keydown', { key: 'Enter', isComposing: true, keyCode: 229 });
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.dispatchEvent('compositionend');
  await input.press('Escape');
  await expect(input).toBeFocused();
  await input.press('Escape');
  await expect(input).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await page.keyboard.press('F2');
  await input.fill('first');
  await input.press('Alt+Enter');
  await input.pressSequentially('second');
  await expect(input).toHaveValue('first\nsecond');
  await input.press('Escape');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
});

test('in-cell formula references highlight the grid and clear on cancel', async ({ page }) => {
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const count = () =>
    page.locator('.os-canvas').evaluate((canvas: HTMLCanvasElement) => {
      const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      let blue = 0;
      for (let i = 0; i < data.length; i += 4)
        if (data[i] === 37 && data[i + 1] === 99 && data[i + 2] === 235) blue++;
      return blue;
    });
  await input.fill('=SUM(B2)');
  await expect.poll(count).toBeGreaterThan(10);
  await input.press('Escape');
  await expect.poll(count).toBe(0);
});

test('in-cell reference drag preserves the target and commits exactly once', async ({ page }) => {
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('D2:D3')
      .setValues([[5], [8]]);
    return book.revision;
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
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.pressSequentially(')');
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before + 1);
  const result = await page.evaluate(async () => {
    const sheet = window.opensheet.getWorkbook().getSheets()[0];
    return {
      target: await sheet.range('A1').getValues(),
      references: await sheet.range('D2:D3').getValues(),
      selection: window.opensheet.selection.get(),
    };
  });
  expect(result.target).toEqual([[13]]);
  expect(result.references).toEqual([[5], [8]]);
  expect(result.selection).toMatchObject({ startRow: 1, startColumn: 0 });
});

test('canceling an in-cell reference drag restores the original selection without a write', async ({
  page,
}) => {
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await input.fill('=SUM(');
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
  await expect(input).toHaveValue('=SUM(B2');
  await input.press('Escape');
  await expect(input).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 0,
    startColumn: 0,
  });
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
});

test('in-cell cross-sheet picking moves the draft to the formula bar without committing', async ({
  page,
}) => {
  const setup = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const origin = book.getSheets()[0];
    const other = await book.addSheet("O'Brien cell");
    await other.range('A1:A2').setValues([[4], [9]]);
    return { revision: book.revision, originId: origin.id, originName: origin.name };
  });
  const cell = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const formula = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await cell.fill('=SUM(');
  await page.getByRole('tab', { name: "O'Brien cell", exact: true }).click();
  await expect(cell).toBeHidden();
  await expect(formula).toBeFocused();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(setup.revision);
  const grid = await page.locator('.os-grid').boundingBox();
  await page.mouse.move(grid!.x + 48 + 8, grid!.y + 30 + 8);
  await page.mouse.down();
  await page.mouse.move(grid!.x + 48 + 8, grid!.y + 30 + 30 + 8, { steps: 3 });
  await page.mouse.up();
  await expect(formula).toHaveValue("=SUM('O''Brien cell'!A1:A2");
  await formula.pressSequentially(')');
  await formula.press('Enter');
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
});

test('Alt+Down transfers an in-cell draft to keyboard reference picking and can cancel safely', async ({
  page,
}) => {
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  const cell = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const formula = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await cell.fill('=SUM(');
  await cell.press('Alt+ArrowDown');
  await expect(cell).toBeHidden();
  await expect(formula).toBeFocused();
  await formula.press('PageDown');
  await formula.press('Shift+ArrowDown');
  await expect(formula).toHaveValue("=SUM('Notes'!A1:A2");
  await formula.press('Escape');
  await expect(formula).toHaveValue('=SUM(');
  await expect(formula).toBeFocused();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await formula.press('Escape');
  await expect(formula).not.toBeFocused();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
});

test('composition cannot be committed by clicking the grid or switching sheets', async ({
  page,
}) => {
  const input = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  const before = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  const origin = await page.evaluate(() => window.opensheet.selection.get()!.sheetId);
  await input.fill('=SUM(');
  await input.dispatchEvent('compositionstart');
  const grid = await page.locator('.os-grid').boundingBox();
  await page.mouse.click(grid!.x + 48 + 3 * 128 + 8, grid!.y + 30 + 30 + 8);
  await expect(input).toBeFocused();
  await expect(input).toHaveValue('=SUM(');
  await page.getByRole('tab', { name: 'Notes', exact: true }).click();
  await expect(input).toBeFocused();
  await page.evaluate(() => {
    const other = window.opensheet.getWorkbook().getSheetByName('Notes')!;
    window.opensheet.selection.set({
      sheetId: other.id,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
  });
  expect(await page.evaluate(() => window.opensheet.selection.get()!.sheetId)).toBe(origin);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.dispatchEvent('compositionend');
  await input.press('Escape');
  await expect(input).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
});

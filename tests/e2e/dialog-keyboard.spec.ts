import { workbookTool } from './workbook-tools.js';
import { expect, test } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 });
  });
});

test('sort levels and actions are keyboard reachable after dynamic changes', async ({ page }) => {
  await workbookTool(page, 'sort-custom');
  await expect(page.getByRole('dialog', { name: 'Custom sort' })).toBeVisible();
  await page.getByLabel('Level 1 order').focus();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Add sort level' })).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(page.getByLabel('Level 2 column')).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByLabel('Level 2 order')).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Remove level 2' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Apply sort' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog', { name: 'Custom sort' })).not.toBeVisible();
});

test('filter condition controls remain keyboard reachable when a condition is added', async ({
  page,
}) => {
  await workbookTool(page, 'filter-custom');
  await expect(page.getByRole('dialog', { name: 'Custom filter' })).toBeVisible();
  await page.getByLabel('Condition 1 operator').focus();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Add condition' })).toBeFocused();
  await page.keyboard.press('Enter');
  await page.getByLabel('Condition 2 operator').focus();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Remove condition 2' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Add condition' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Apply filter' })).toBeFocused();
});

for (const kind of ['sort', 'filter'] as const) {
  test(`${kind} adds a row without replacing existing controls and can remove the first row`, async ({
    page,
  }) => {
    await workbookTool(page, kind === 'sort' ? 'sort-custom' : 'filter-custom');
    const firstLabel = kind === 'sort' ? 'Level 1 column' : 'Condition 1 column';
    const addLabel = kind === 'sort' ? 'Add sort level' : 'Add condition';
    const removeLabel = kind === 'sort' ? 'Remove level 1' : 'Remove condition 1';
    await page.getByLabel(firstLabel).evaluate((element) => {
      (window as any).__originalDialogControl = element;
    });
    await page.getByRole('button', { name: addLabel }).click();
    expect(
      await page
        .getByLabel(firstLabel)
        .evaluate((element) => (window as any).__originalDialogControl === element),
    ).toBe(true);
    await expect(page.getByRole('button', { name: removeLabel })).toBeEnabled();
    await page.getByRole('button', { name: removeLabel }).click();
    await expect(page.getByLabel(firstLabel)).toBeVisible();
    await expect(page.getByRole('button', { name: removeLabel })).toBeDisabled();
    if (kind === 'sort') await expect(page.getByLabel(firstLabel)).toHaveValue('1');
    await page.keyboard.press('Escape');
  });

  test(`${kind} retains and renumbers surviving controls after repeated removals`, async ({
    page,
  }) => {
    await page.evaluate(() => {
      const app = window.opensheet;
      const sheet = app.getWorkbook().getSheets()[0];
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 4,
        startColumn: 0,
        endColumn: 3,
      });
    });
    await workbookTool(page, kind === 'sort' ? 'sort-custom' : 'filter-custom');
    const add = page.getByRole('button', {
      name: kind === 'sort' ? 'Add sort level' : 'Add condition',
    });
    await add.click();
    await add.click();
    const second = page.getByLabel(kind === 'sort' ? 'Level 2 column' : 'Condition 2 column');
    if (kind === 'filter') {
      await page.getByLabel('Condition 2 operator').selectOption('eq');
      await page.getByLabel('Condition 2 value', { exact: true }).fill('kept');
    }
    await second.evaluate((element) => {
      (window as any).__survivingDialogControl = element;
    });
    await page
      .getByRole('button', { name: kind === 'sort' ? 'Remove level 1' : 'Remove condition 1' })
      .click();
    const first = page.getByLabel(kind === 'sort' ? 'Level 1 column' : 'Condition 1 column');
    expect(
      await first.evaluate((element) => (window as any).__survivingDialogControl === element),
    ).toBe(true);
    await page
      .getByRole('button', { name: kind === 'sort' ? 'Remove level 2' : 'Remove condition 2' })
      .click();
    await expect(
      page.getByRole('button', { name: kind === 'sort' ? 'Remove level 1' : 'Remove condition 1' }),
    ).toBeDisabled();
    if (kind === 'sort') {
      await expect(first).toHaveValue('1');
      await page.getByLabel('Level 1 order').selectOption('desc');
    } else {
      await expect(page.getByLabel('Condition 1 value', { exact: true })).toHaveValue('kept');
      await page.getByLabel('Condition 1 value', { exact: true }).fill('changed');
    }
    await page
      .getByRole('button', { name: kind === 'sort' ? 'Apply sort' : 'Apply filter' })
      .click();
    await expect(page.getByRole('dialog')).not.toBeVisible();
    if (kind === 'filter') {
      const filter = await page.evaluate(() => {
        const book = window.opensheet.getWorkbook();
        return book.sheetData(book.getSheets()[0].id).filter;
      });
      expect(filter).toMatchObject({ criteria: [{ operator: 'eq', value: 'changed' }] });
    }
  });
}

test('series direction and submit are keyboard reachable', async ({ page }) => {
  await workbookTool(page, 'fill-series');
  await expect(page.getByLabel('Step', { exact: true })).toBeEditable();
  await page.getByLabel('Step', { exact: true }).focus();
  await page.keyboard.press('Tab');
  await expect(page.getByLabel('Direction')).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Fill series' })).toBeFocused();
});

test('paste action is reachable after keyboard clipboard capture', async ({ page }) => {
  await workbookTool(page, 'paste-special');
  const box = page.getByLabel('Paste clipboard here');
  await box.evaluate((element) => {
    const clipboard = new DataTransfer();
    clipboard.setData('text/plain', 'hello');
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboard });
    element.dispatchEvent(event);
  });
  await box.focus();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Apply paste' })).toBeFocused();
});

test('note link and actions are keyboard reachable', async ({ page }) => {
  await workbookTool(page, 'annotation');
  await expect(page.getByLabel('Link URL')).toBeEditable();
  await page.getByLabel('Link URL').fill('https://example.com/');
  await page.getByLabel('Link description').focus();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('link', { name: 'Open link' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Save note and link' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Remove note and link' })).toBeFocused();
});

import { expect, test } from '@playwright/test';

test('explains a selected formula error without changing the workbook or a draft', async ({
  page,
}) => {
  await page.goto('/#budget');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => window.opensheet.ready());
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book.getSheets()[0].range('A1').setInput('=1/0');
    return book.revision;
  });
  const button = page.getByRole('button', { name: 'Explain formula error #DIV/0!' });
  await expect(button).toBeVisible();
  await button.focus();
  await button.press('Enter');
  await expect(button).toHaveAttribute('aria-expanded', 'true');
  await expect(page.locator('.os-formula-error-details')).toContainText('referenced cell');
  await button.press('Escape');
  await expect(button).toHaveAttribute('aria-expanded', 'false');
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await input.fill('=2+2');
  await expect(button).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(before + 1);
  await expect(button).toBeHidden();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[4]]);
});

test('follows propagated errors and clears stale explanations after recalculation and selection', async ({
  page,
}) => {
  await page.goto('/#budget');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const sheet = window.opensheet.getWorkbook().getSheets()[0];
    await sheet.range('A1').setInput('=1/0');
    await sheet.range('B1').setInput('=A1');
    const selection = window.opensheet.selection.get()!;
    window.opensheet.selection.set({ ...selection, startColumn: 1, endColumn: 2 });
  });
  const button = page.getByRole('button', { name: 'Explain formula error #DIV/0!' });
  await expect(button).toBeVisible();
  await button.focus();
  await button.press('Enter');
  await expect(page.locator('.os-formula-error-details')).toContainText(
    'may originate in a referenced cell',
  );
  await page.evaluate(async () => {
    await window.opensheet.getWorkbook().getSheets()[0].range('A1').setInput('=1/1');
  });
  await expect(button).toBeHidden();
  await expect(page.locator('.os-formula-error-details')).toBeHidden();
  await expect(page.getByRole('grid', { name: 'Spreadsheet' })).toBeFocused();
  await page.evaluate(async () => {
    const sheet = window.opensheet.getWorkbook().getSheets()[0];
    await sheet.range('A1').setInput('=1/0');
  });
  await expect(button).toBeVisible();
  await button.focus();
  await button.press('Enter');
  await page.evaluate(() => {
    const selection = window.opensheet.selection.get()!;
    window.opensheet.selection.set({ ...selection, startColumn: 2, endColumn: 3 });
  });
  await expect(button).toBeHidden();
  await expect(page.locator('.os-formula-error-details')).toBeHidden();
  await expect(page.getByRole('grid', { name: 'Spreadsheet' })).toBeFocused();
});

test('keeps error guidance inside a narrow editor and preserves focus moved elsewhere', async ({
  page,
}) => {
  await page.setViewportSize({ width: 700, height: 650 });
  await page.goto('/#budget');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet.getWorkbook().getSheets()[0].range('A1').setInput('=1/0');
    const bar = document.querySelector<HTMLElement>('.os-formula-bar')!;
    bar.style.width = '260px';
  });
  const button = page.getByRole('button', { name: 'Explain formula error #DIV/0!' });
  await button.click();
  const bounds = await page.locator('.os-formula-error-details').evaluate((element) => {
    const box = element.getBoundingClientRect();
    const bar = element.parentElement!.getBoundingClientRect();
    return { left: box.left, right: box.right, barLeft: bar.left, barRight: bar.right };
  });
  expect(bounds.left).toBeGreaterThanOrEqual(bounds.barLeft);
  expect(bounds.right).toBeLessThanOrEqual(bounds.barRight);
  const address = page.getByRole('textbox', { name: 'Cell address' });
  await address.focus();
  await page.evaluate(async () => {
    await window.opensheet.getWorkbook().getSheets()[0].range('A1').setInput('=1/1');
  });
  await expect(button).toBeHidden();
  await expect(address).toBeFocused();
});

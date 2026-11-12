import { expect, test } from '@playwright/test';

for (const mode of ['bar', 'cell'] as const) {
  test(`${mode} editor cycles a reference after the highlighting limit`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
    if (mode === 'cell') {
      await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
      await page.keyboard.press('F2');
    }
    const input =
      mode === 'bar'
        ? page.getByRole('combobox', { name: 'Formula bar' })
        : page.getByRole('textbox', { name: 'Edit cell', exact: true });
    const formula = `=SUM(${Array.from({ length: 80 }, () => 'A1').join(',')},B2)`;
    await input.fill(formula);
    const target = formula.lastIndexOf('B2');
    await input.evaluate(
      (element: HTMLInputElement, position) => element.setSelectionRange(position, position),
      target + 1,
    );
    await input.press('F4');
    await expect(input).toHaveValue(formula.slice(0, target) + '$B$2)');
    expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
    await input.press('Escape');
  });
  test(`${mode} editor cycles reference anchors without committing the draft`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
    if (mode === 'cell') {
      await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
      await page.keyboard.press('F2');
    }
    const input =
      mode === 'bar'
        ? page.getByRole('combobox', { name: 'Formula bar' })
        : page.getByRole('textbox', { name: 'Edit cell', exact: true });
    await input.fill('=SUM(B2:C3)');
    await input.evaluate((element: HTMLInputElement) => element.setSelectionRange(5, 10));
    await input.press('F4');
    await expect(input).toHaveValue('=SUM($B$2:$C$3)');
    await input.press('F4');
    await expect(input).toHaveValue('=SUM(B$2:C$3)');
    await input.press('F4');
    await expect(input).toHaveValue('=SUM($B2:$C3)');
    await input.press('F4');
    await expect(input).toHaveValue('=SUM(B2:C3)');
    await input.dispatchEvent('compositionstart');
    await input.dispatchEvent('keydown', { key: 'F4', isComposing: true, keyCode: 229 });
    await expect(input).toHaveValue('=SUM(B2:C3)');
    await input.dispatchEvent('compositionend');
    expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
    await input.press('F4');
    await input.press('Enter');
    await expect
      .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
      .toBe(revision + 1);
    const result = await page.evaluate(async () => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      const formula = (
        await book.readCells(
          sheet.id,
          { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
          { fields: ['input'] },
        )
      ).cells[0].input;
      await book.undo();
      const undone = (
        await book.readCells(
          sheet.id,
          { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
          { fields: ['input'] },
        )
      ).cells[0].input;
      await book.redo();
      await app.open(book.id);
      const reopened = app.getWorkbook();
      const restored = (
        await reopened.readCells(
          reopened.getSheets()[0].id,
          { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
          { fields: ['input'] },
        )
      ).cells[0].input;
      return { formula, undone, restored };
    });
    expect(result.formula).toEqual({ type: 'formula', expression: 'SUM($B$2:$C$3)' });
    expect(result.undone).not.toEqual(result.formula);
    expect(result.restored).toEqual(result.formula);
  });
}

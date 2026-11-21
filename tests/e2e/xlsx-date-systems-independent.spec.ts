import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';

for (const system of ['1900', '1904'] as const) {
  test(`independent ${system} dates import, recalculate and reopen`, async ({ page }) => {
    const root = new URL('../../benchmarks/date-system-openpyxl-2026-11-22/', import.meta.url);
    const bytes = readFileSync(new URL(`dates-${system}.xlsx`, root));
    const evidence = JSON.parse(readFileSync(new URL(`dates-${system}.xlsx.json`, root), 'utf8'));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sha256);
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    await page.locator('#file').setInputFiles({
      name: `dates-${system}.xlsx`,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: bytes,
    });
    const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
    await expect(dialog).toBeVisible();
    await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
    await expect(page.getByRole('tab', { name: 'Dates', exact: true })).toBeVisible();
    const result = await page.evaluate(async () => {
      const app = window.opensheet;
      let book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      const read = async () => ({
        system: (await book.toJSON()).dateSystem,
        values: await sheet.range('A1:A3').getValues(),
        display: await sheet.range('A1:A3').getDisplayValues(),
        formula: await sheet.range('B1').getFormulas(),
        calculated: await sheet.range('B1').getValues(),
      });
      const initial = await read();
      await sheet.range('A1').setValues([[(initial.values[0][0] as number) + 1]]);
      const changed = await read();
      await book.undo();
      const undone = await read();
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      const reopened = await read();
      return { initial, changed, undone, reopened };
    });
    expect(result.initial.system).toBe(system);
    expect(result.initial.values).toEqual(evidence.xlsxSerials.map((value: number) => [value]));
    expect(result.initial.display).toEqual([['1904-01-02'], ['2024-02-29 12:30'], ['2100-03-01']]);
    expect(result.initial.formula).toEqual([['A2-A1']]);
    expect(result.initial.calculated[0][0]).toBeCloseTo(
      evidence.xlsxSerials[1] - evidence.xlsxSerials[0],
      8,
    );
    expect(result.changed.values[0][0]).toBe(evidence.xlsxSerials[0] + 1);
    expect(result.changed.display).toEqual([['1904-01-03'], ['2024-02-29 12:30'], ['2100-03-01']]);
    expect(result.changed.calculated[0][0]).toBeCloseTo(
      (result.initial.calculated[0][0] as number) - 1,
      8,
    );
    expect(result.undone).toEqual(result.initial);
    expect(result.reopened).toEqual(result.changed);
  });
}

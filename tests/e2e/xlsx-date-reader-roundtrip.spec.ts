import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';

const evidence = JSON.parse(
  readFileSync('benchmarks/date-system-reader-verified-2026-11-22.json', 'utf8'),
);
for (const system of ['1900', '1904'] as const) {
  for (const index of [2, 3] as const) {
    test(`${system} independent ${index === 2 ? 'control' : 'OpenSheet round-trip'} recalculates and reopens`, async ({
      page,
    }) => {
      const entry = evidence.systems.find((item: { system: string }) => item.system === system);
      const file = entry.files[index],
        observed = entry.independent.files[index];
      const bytes = readFileSync(file.path);
      expect(createHash('sha256').update(bytes).digest('hex')).toBe(file.sha256);
      const expected = observed.rawSerials.map(Number);
      await page.goto('/#budget');
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      await page.locator('#file').setInputFiles({
        name: 'dates.xlsx',
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
        return { initial, changed, undone, reopened: await read() };
      });
      expect(result.initial.system).toBe(system);
      expect(result.initial.values).toEqual(expected.map((value: number) => [value]));
      expect(result.initial.display).toEqual([
        ['1904-01-02'],
        ['2024-02-29 12:30'],
        ['2100-03-01'],
      ]);
      expect(result.initial.formula).toEqual([['A2-A1']]);
      expect(result.initial.calculated[0][0]).toBeCloseTo(observed.cached, 8);
      expect(result.changed.calculated[0][0]).toBeCloseTo(observed.cached - 1, 8);
      expect(result.changed.display[0]).toEqual(['1904-01-03']);
      expect(result.undone).toEqual(result.initial);
      expect(result.reopened).toEqual(result.changed);
    });
  }
}

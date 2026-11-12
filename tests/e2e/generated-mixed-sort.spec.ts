import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('seeded mixed-type sorts preserve row payloads and history', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Mixed', rows: 10, columns: 2 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    type Value = number | string | boolean | null | { error: string };
    type Row = { value: Value; id: string };
    const original: Row[] = [
      { value: 2, id: 'n2' },
      { value: -1, id: 'negative' },
      { value: 'apple', id: 'lower' },
      { value: 'Apple', id: 'upper' },
      { value: '10', id: 'text-number' },
      { value: false, id: 'false' },
      { value: true, id: 'true' },
      { value: { error: '#DIV/0!' }, id: 'error' },
      { value: null, id: 'blank1' },
      { value: null, id: 'blank2' },
    ];
    for (const seed of [17, 241, 1009]) {
      let rng = seed;
      const random = (n: number) => {
        rng = (Math.imul(rng, 1664525) + 1013904223) >>> 0;
        return Math.floor((rng / 0x100000000) * n);
      };
      const rows = original.slice();
      for (let i = rows.length - 1; i > 0; i--) {
        const j = random(i + 1);
        [rows[i], rows[j]] = [rows[j], rows[i]];
      }
      for (const direction of ['asc', 'desc'] as const) {
        const app = window.opensheet;
        await app.load(JSON.parse(fileJSON));
        let book = app.getWorkbook(),
          sheet = book.getSheets()[0];
        await sheet
          .range('A1:B10')
          .setValues(
            rows.map(({ value, id }) => [
              typeof value === 'object' && value !== null ? null : value,
              id,
            ]),
          );
        await sheet
          .range(`A${rows.findIndex((row) => row.id === 'error') + 1}`)
          .setFormulas([['1/0']]);
        // Fixed semantic groups give case-insensitive ties the input order in either direction.
        const groups =
          direction === 'asc'
            ? [
                ['negative'],
                ['n2'],
                ['text-number'],
                ['lower', 'upper'],
                ['false'],
                ['true'],
                ['error'],
                ['blank1', 'blank2'],
              ]
            : [
                ['error'],
                ['true'],
                ['false'],
                ['lower', 'upper'],
                ['text-number'],
                ['n2'],
                ['negative'],
                ['blank1', 'blank2'],
              ];
        const expected = groups.flatMap((group) =>
          rows.filter((row) => group.includes(row.id)).map(({ value, id }) => [value, id]),
        );
        const before = rows.map(({ value, id }) => [value, id]);
        await sheet.sort('A1:B10', 0, direction);
        let actual = await sheet.range('A1:B10').getValues();
        if (JSON.stringify(actual) !== JSON.stringify(expected))
          return { seed, direction, stage: 'sort', expected, actual };
        await app.open(book.id);
        book = app.getWorkbook();
        sheet = book.getSheets()[0];
        if (!(await book.undo())) return { seed, direction, stage: 'undo failed' };
        actual = await sheet.range('A1:B10').getValues();
        if (JSON.stringify(actual) !== JSON.stringify(before))
          return { seed, direction, stage: 'undo', before, actual };
        if (!(await book.redo())) return { seed, direction, stage: 'redo failed' };
        actual = await sheet.range('A1:B10').getValues();
        if (JSON.stringify(actual) !== JSON.stringify(expected))
          return { seed, direction, stage: 'redo', expected, actual };
      }
    }
    return null;
  }, file);
  expect(result).toBeNull();
});

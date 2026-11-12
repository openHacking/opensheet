import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('seeded filter, sort and edit history agree with independent row visibility after reopen', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Filter model', rows: 8, columns: 2 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    for (const seed of [223, 901, 8191]) {
      await app.load(JSON.parse(fileJSON));
      const ids = Array.from({ length: 8 }, (_, index) => `row-${index}`);
      await app
        .getWorkbook()
        .getSheets()[0]
        .range('B1:B8')
        .setValues(ids.map((id) => [id]));
      type Model = { values: (number | null)[]; ids: string[]; threshold: number | null };
      let model: Model = { values: Array(8).fill(null), ids, threshold: null };
      let past: Model[] = [],
        future: Model[] = [];
      const copy = (value: Model): Model => ({
        values: value.values.slice(),
        ids: value.ids.slice(),
        threshold: value.threshold,
      });
      let rng = seed;
      const random = (n: number) => {
        rng = (Math.imul(rng, 1664525) + 1013904223) >>> 0;
        return Math.floor((rng / 0x100000000) * n);
      };
      const trace: string[] = [],
        covered = new Set<string>();
      for (let step = 0; step < 36; step++) {
        let book = app.getWorkbook(),
          sheet = book.getSheets()[0];
        const action = step < 11 ? [0, 1, 0, 3, 4, 2, 3, 1, 5, 3, 4][step] : random(6);
        if (action === 0) {
          const row = random(8),
            value = step + seed;
          past.push(copy(model));
          future = [];
          await sheet
            .range({ startRow: row, endRow: row + 1, startColumn: 0, endColumn: 1 })
            .setValues([[value]]);
          model.values[row] = value;
          trace.push('edit');
        } else if (action === 1) {
          const threshold = seed + random(30);
          past.push(copy(model));
          future = [];
          await sheet.setFilter({
            mode: 'all',
            headerRows: 0,
            criteria: [{ column: 0, operator: 'gt', value: threshold }],
          });
          model.threshold = threshold;
          trace.push('filter');
        } else if (action === 2 && model.threshold !== null) {
          past.push(copy(model));
          future = [];
          await sheet.setFilter(null);
          model.threshold = null;
          trace.push('clear');
        } else if (action === 5) {
          if (model.threshold !== null) covered.add('sort-filtered');
          const direction = random(2) ? 'asc' : 'desc';
          past.push(copy(model));
          future = [];
          await sheet.sort('A1:B8', 0, direction);
          const sorted = model.values
            .map((value, index) => ({ value, id: model.ids[index], index }))
            .sort((a, b) => {
              if (a.value === null && b.value === null) return a.index - b.index;
              if (a.value === null) return 1;
              if (b.value === null) return -1;
              return (
                (direction === 'asc' ? a.value - b.value : b.value - a.value) || a.index - b.index
              );
            });
          model.values = sorted.map(({ value }) => value);
          model.ids = sorted.map(({ id }) => id);
          trace.push('sort');
        } else if (action === 3 && past.length) {
          future.push(copy(model));
          model = past.pop()!;
          if (!(await book.undo())) return { seed, step, trace, error: 'undo failed' };
          trace.push('undo');
        } else if (action === 4 && future.length) {
          past.push(copy(model));
          model = future.pop()!;
          if (!(await book.redo())) return { seed, step, trace, error: 'redo failed' };
          trace.push('redo');
        } else trace.push('skip');
        covered.add(trace[trace.length - 1]);
        if (step === 3 || step % 9 === 8) {
          await app.open(book.id);
          book = app.getWorkbook();
          sheet = book.getSheets()[0];
          trace.push('reopen');
        }
        const actual = await sheet.range('A1:B8').getValues();
        const hidden = book.filteredRows(sheet.id);
        const expectedHidden =
          model.threshold === null
            ? []
            : model.values.flatMap((value, index) =>
                value === null || value <= model.threshold! ? [index] : [],
              );
        if (
          JSON.stringify(actual.map((row) => row[0])) !== JSON.stringify(model.values) ||
          JSON.stringify(actual.map((row) => row[1])) !== JSON.stringify(model.ids) ||
          JSON.stringify(hidden) !== JSON.stringify(expectedHidden)
        )
          return { seed, step, trace, model, actual, hidden, expectedHidden };
      }
      for (const action of ['edit', 'filter', 'clear', 'sort', 'sort-filtered', 'undo', 'redo'])
        if (!covered.has(action)) return { seed, trace, error: `missing ${action}` };
    }
    return null;
  }, file);
  expect(result).toBeNull();
});

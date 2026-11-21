import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';

const fixture = readFileSync(
  new URL('../../apps/playground/src/budget-fixture.ts', import.meta.url),
);

test('budget workload recalculates cross-sheet mixed references through history and cold reopen', async ({
  page,
}) => {
  expect(createHash('sha256').update(fixture).digest('hex')).toBe(
    'ffe1528cd788ae9c55e3e566a8fccf3e034445ce73c0f358851c362b28be52dd',
  );
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const existing = window.opensheet.getWorkbook();
    const Constructor = existing.constructor as new (
      snapshot: any,
      options: object,
      id?: string,
    ) => typeof existing;
    const snapshot = await existing.toJSON();
    snapshot.workbookId = `budget-workload-${crypto.randomUUID()}`;
    const options = { database: `budget-workload-${crypto.randomUUID()}`, cacheBytes: 0 };
    let book = await new Constructor(snapshot, options).ready();
    try {
      let [budget, notes] = book.getSheets();
      await notes
        .range('D2:E2')
        .setFormulas([
          ["'Launch budget'!$D2*'Launch budget'!E$2", "'Launch budget'!$D$2+'Launch budget'!E$2"],
        ]);
      const read = async () => ({
        budget: await budget.range('D2:F2').getValues(),
        total: await budget.range('F11').getValues(),
        references: await notes.range('D2:E2').getValues(),
        formulas: await notes.range('D2:E2').getFormulas(),
      });
      const initial = await read();
      await budget.range('D2').setValues([[3]]);
      const edited = await read();
      await book.undo();
      const undone = await read();
      await book.redo();
      const redone = await read();
      const id = book.id;
      const revision = book.revision;
      await book.close();
      book = await new Constructor(undefined, options, id).ready();
      [budget, notes] = book.getSheets();
      const reopened = await read();
      const cacheBytes = (await book.storageStats()).cacheBytes;
      return {
        initial,
        edited,
        undone,
        redone,
        reopened,
        revision,
        reopenedRevision: book.revision,
        cacheBytes,
      };
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
  const formulas = [
    ["'Launch budget'!$D2*'Launch budget'!E$2", "'Launch budget'!$D$2+'Launch budget'!E$2"],
  ];
  const initial = {
    budget: [[1, 2400, 2400]],
    total: [[13097]],
    references: [[2400, 2401]],
    formulas,
  };
  const edited = {
    budget: [[3, 2400, 7200]],
    total: [[17897]],
    references: [[7200, 2403]],
    formulas,
  };
  expect(result.initial).toEqual(initial);
  expect(result.edited).toEqual(edited);
  expect(result.undone).toEqual(initial);
  expect(result.redone).toEqual(edited);
  expect(result.reopened).toEqual(edited);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.cacheBytes).toBe(0);
});

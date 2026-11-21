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

test('budget calculations survive verified Worker page-cache eviction and reopen', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let budgetReadWatch = false, budgetPageReads = [];
addEventListener('message', ({data}) => {
  if (data.budgetReadStart) { budgetReadWatch = true; budgetPageReads = []; }
  if (data.budgetReadStop) { budgetReadWatch = false; postMessage({budgetPageReads}); }
});
const budgetOriginalGet = IDBObjectStore.prototype.get;
IDBObjectStore.prototype.get = function(...args) {
  if (budgetReadWatch && this.name === 'records') budgetPageReads.push(String(args[0]));
  return budgetOriginalGet.apply(this, args);
};`,
    });
  });
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
    snapshot.workbookId = `budget-eviction-${crypto.randomUUID()}`;
    const options = { database: `budget-eviction-${crypto.randomUUID()}`, cacheBytes: 64 * 1024 };
    let book = await new Constructor(snapshot, options).ready();
    try {
      let [budget, notes] = book.getSheets();
      await notes.range('D2').setFormulas([["'Launch budget'!$D2*'Launch budget'!E$2"]]);
      const filler = await book.addSheet('Cache filler', { rows: 1024, columns: 1 });
      await filler
        .range('A1:A1024')
        .setValues(Array.from({ length: 1024 }, (_, i) => [`${i}:` + 'x'.repeat(512)]));
      const id = book.id;
      await book.close();
      book = await new Constructor(undefined, options, id).ready();
      [budget, notes] = book.getSheets();
      const read = async () => ({
        budget: await budget.range('D2:F2').getValues(),
        reference: await notes.range('D2').getValues(),
      });
      const measure = async (operation: () => Promise<unknown>) => {
        const worker = (book as any).worker as Worker;
        worker.postMessage({ budgetReadStart: true });
        const value = await operation();
        const reads = await new Promise<string[]>((resolve) => {
          worker.addEventListener('message', function receive({ data }) {
            if (data.budgetPageReads) {
              worker.removeEventListener('message', receive);
              resolve(data.budgetPageReads);
            }
          });
          worker.postMessage({ budgetReadStop: true });
        });
        return { value, reads };
      };
      const initial = await measure(read);
      const warm = await measure(read);
      const revision = book.revision;
      const values = await book.getSheets()[2].range('A1:A1024').getValues();
      const fillerCorrect = values.every((row, i) => row[0] === `${i}:` + 'x'.repeat(512));
      const afterEviction = await measure(read);
      const afterReadRevision = book.revision;
      await budget.range('D2').setValues([[3]]);
      const edited = await read();
      await book.undo();
      const undone = await read();
      await book.redo();
      const redone = await read();
      await book.close();
      book = await new Constructor(undefined, options, id).ready();
      [budget, notes] = book.getSheets();
      const reopened = await read();
      return {
        initial,
        warm,
        afterEviction,
        revision,
        afterReadRevision,
        fillerCorrect,
        edited,
        undone,
        redone,
        reopened,
      };
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
  const initial = { budget: [[1, 2400, 2400]], reference: [[2400]] };
  const edited = { budget: [[3, 2400, 7200]], reference: [[7200]] };
  expect(result.initial.value).toEqual(initial);
  expect(result.warm.value).toEqual(initial);
  expect(result.afterEviction.value).toEqual(initial);
  expect(result.initial.reads.length).toBeGreaterThan(0);
  expect(result.warm.reads).toEqual([]);
  expect(result.afterEviction.reads.some((id) => result.initial.reads.includes(id))).toBe(true);
  expect(result.fillerCorrect).toBe(true);
  expect(result.afterReadRevision).toBe(result.revision);
  expect(result.edited).toEqual(edited);
  expect(result.undone).toEqual(initial);
  expect(result.redone).toEqual(edited);
  expect(result.reopened).toEqual(edited);
});

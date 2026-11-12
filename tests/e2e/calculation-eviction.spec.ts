import { expect, test } from '@playwright/test';

test('small calculation cache cannot retain stale transitive results after eviction', async ({
  page,
}) => {
  await page.goto('/#budget');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const values = await page.evaluate(async () => {
    const existing = window.opensheet.getWorkbook();
    const Constructor = existing.constructor as new (
      snapshot: any,
      options: object,
      id?: string,
    ) => typeof existing;
    const snapshot = await existing.toJSON();
    snapshot.workbookId = `eviction-${crypto.randomUUID()}`;
    const options = { database: `eviction-${crypto.randomUUID()}`, cacheBytes: 3200 };
    let book = await new Constructor(snapshot, options).ready();
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[1]]);
      await sheet.range('B1:C1').setFormulas([['A1+1', 'B1+1']]);
      const before = (await sheet.range('C1').getValues())[0][0];
      await sheet.range('A1').setValues([[7]]);
      const after = (await sheet.range('C1').getValues())[0][0];
      await book.undo();
      const undone = (await sheet.range('C1').getValues())[0][0];
      await book.redo();
      const redone = (await sheet.range('C1').getValues())[0][0];
      const id = book.id;
      await book.close();
      book = await new Constructor(undefined, options, id).ready();
      sheet = book.getSheets()[0];
      const reopened = (await sheet.range('C1').getValues())[0][0];
      return { before, after, undone, redone, reopened };
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
  expect(values).toEqual({ before: 3, after: 9, undone: 3, redone: 9, reopened: 9 });
});

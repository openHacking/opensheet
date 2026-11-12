import { expect, test } from '@playwright/test';

test('sparse value and formula matrices fail before dispatch without partial edits', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 2, columns: 2 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:B2').setValues([
      [1, 2],
      [3, 4],
    ]);
    const revision = book.revision;
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    let writes = 0;
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') writes++;
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    const missingRow = [
      [9, 9],
      [9, 9],
    ];
    delete missingRow[1];
    const missingCell = [
      [9, 9],
      [9, 9],
    ];
    delete missingCell[1][1];
    const formulaRow = [
      ['1', '2'],
      ['3', '4'],
    ];
    delete formulaRow[1];
    const formulaCell = [
      ['1', '2'],
      ['3', '4'],
    ];
    delete formulaCell[1][1];
    const operations = [
      () => sheet.range('A1:B2').setValues(missingRow),
      () => sheet.range('A1:B2').setValues(missingCell),
      () => sheet.range('A1:B2').setFormulas(formulaRow),
      () => sheet.range('A1:B2').setFormulas(formulaCell),
    ];
    const errors = [];
    for (const operation of operations) {
      try {
        await operation();
        errors.push('success');
      } catch (error) {
        errors.push((error as { code?: string }).code);
      }
    }
    worker.postMessage = post;
    const after = { revision: book.revision, values: await sheet.range('A1:B2').getValues() };
    await sheet.range('A1:B2').setValues([
      [null, 6],
      [7, 8],
    ]);
    const explicitBlank = await sheet.range('A1:B2').getValues();
    await book.undo();
    const undone = await sheet.range('A1:B2').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      revision,
      errors,
      writes,
      after,
      explicitBlank,
      undone,
      reopened: await sheet.range('A1:B2').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.errors).toEqual(Array(4).fill('INVALID_ARGUMENT'));
  expect(result.writes).toBe(0);
  expect(result.after).toEqual({
    revision: result.revision,
    values: [
      [1, 2],
      [3, 4],
    ],
  });
  expect(result.explicitBlank).toEqual([
    [null, 6],
    [7, 8],
  ]);
  expect(result.undone).toEqual(result.after.values);
  expect(result.reopened).toEqual(result.explicitBlank);
});

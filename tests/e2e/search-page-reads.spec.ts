import { expect, test } from '@playwright/test';

test('records zero-cache page reads for a deep search in a real Worker', async ({ page }) => {
  test.setTimeout(60000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchSearchReads = false, searchReads = [];
addEventListener('message', ({data}) => {
  if (data.startSearchReadCount) { searchReads = []; watchSearchReads = true; }
  if (data.reportSearchReadCount) { watchSearchReads = false; postMessage({searchReadCounts: {total: searchReads.length, unique: new Set(searchReads).size}}); }
});

const originalSearchGet = IDBObjectStore.prototype.get;
IDBObjectStore.prototype.get = function(...args) {
  if (watchSearchReads && this.name === 'records') searchReads.push(String(args[0]));
  return originalSearchGet.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
    let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 2 }] });
    let sheet = book.getSheets()[0];
    await sheet
      .range('A1:A129')
      .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 2}+1`]));
    await sheet.range('A130').setValues([[1]]);
    await sheet.range('B1').setValues([['#LIMIT!']]);
    const id = book.id;
    await book.close();
    book = await openWorkbook(id, { cacheBytes: 0 });
    sheet = book.getSheets()[0];
    const worker = (book as any).worker as Worker;
    worker.postMessage({ startSearchReadCount: true });
    const start = performance.now();
    const codes: string[] = [];
    for (const direction of ['next', 'previous'] as const)
      codes.push(
        await book.search(sheet.id, '#LIMIT!', -1, { direction }).then(
          () => 'success',
          (error: { code: string }) => error.code,
        ),
      );
    const elapsedMs = performance.now() - start;
    const counts = await new Promise<{ total: number; unique: number }>((resolve) => {
      worker.addEventListener('message', function receive({ data }) {
        if (data.searchReadCounts) {
          worker.removeEventListener('message', receive);
          resolve(data.searchReadCounts);
        }
      });
      worker.postMessage({ reportSearchReadCount: true });
    });
    await book.deleteStorage();
    return { codes, elapsedMs, counts };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  console.log('zero-cache search page reads', JSON.stringify(result));
  expect(result.codes).toEqual(['LIMIT_EXCEEDED', 'LIMIT_EXCEEDED']);
  expect(result.counts.unique).toBeGreaterThan(0);
  expect(result.counts.total).toBeGreaterThan(result.counts.unique);
  expect(result.counts.total).toBeLessThan(400);
});

test('zero-cache formula search observes queued edits, undo and reopen', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
    let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 2 }] });
    let sheet = book.getSheets()[0];
    await sheet
      .range('A1:A129')
      .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 2}+1`]));
    await sheet.range('A130').setValues([[1]]);
    const id = book.id;
    await book.close();
    book = await openWorkbook(id, { cacheBytes: 0 });
    sheet = book.getSheets()[0];
    const first = book.search(sheet.id, '2');
    const edit = sheet.range('A130').setValues([[9]]);
    const firstCode = await first.then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    await edit;
    const secondCode = await book.search(sheet.id, '2').then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    await sheet.range('A1:A2').clear();
    const afterClear = await book.search(sheet.id, '10');
    await book.undo();
    const afterUndo = await book.search(sheet.id, '2').then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    await book.close();
    book = await openWorkbook(id, { cacheBytes: 0 });
    const afterReopen = await book.search(book.getSheets()[0].id, '2').then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    const last = await book.getSheets()[0].range('A130').getValues();
    await book.deleteStorage();
    return { firstCode, secondCode, afterClear, afterUndo, afterReopen, last };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.firstCode).toBe('LIMIT_EXCEEDED');
  expect(result.secondCode).toBe('LIMIT_EXCEEDED');
  expect(result.afterClear.count).toBeGreaterThan(0);
  expect(result.afterUndo).toBe('LIMIT_EXCEEDED');
  expect(result.afterReopen).toBe('LIMIT_EXCEEDED');
  expect(result.last).toEqual([[9]]);
});

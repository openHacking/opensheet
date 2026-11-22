import { expect, test } from '@playwright/test';

test('cancelled copy read stops a cached-reference formula at its next yield', async ({ page }) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let armFormulaTimer = false, heldFormulaTimer, formulaTimers = 0;
const originalFormulaTimeout = globalThis.setTimeout;
globalThis.setTimeout = (callback, delay, ...args) => {
  if (armFormulaTimer && delay === 0) {
    formulaTimers++;
    if (formulaTimers === 1) return originalFormulaTimeout(() => { heldFormulaTimer = () => callback(...args); postMessage({formulaCopyStalled: true}); }, 0);
  }
  return originalFormulaTimeout(callback, delay, ...args);
};
addEventListener('message', ({data}) => {
  if (data.armFormulaCopy) { armFormulaTimer = true; formulaTimers = 0; postMessage({formulaCopyArmed: true}); }
  if (data.method === 'cancelRequest' && data.args[0] > 0) postMessage({formulaCopyCancelled: true});
  if (data.releaseFormulaCopy && heldFormulaTimer) { const fn = heldFormulaTimer; heldFormulaTimer = undefined; fn(); }
  if (data.reportFormulaCopy) { armFormulaTimer = false; postMessage({formulaCopyTimers: formulaTimers}); }
});`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
    let book = await createWorkbook(
      { sheets: [{ name: 'Data', rows: 10, columns: 2 }] },
      { cacheBytes: 0 },
    );
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1:A10').setValues(Array.from({ length: 10 }, () => [1]));
      const expression = `IFERROR(SUM(${Array.from({ length: 255 }, () => 'A1:A10').join(',')}),0)`;
      await sheet.range('B1').setFormulas([[expression]]);
      const worker = (book as any).worker as Worker;
      const wait = (field: string) =>
        new Promise<any>((resolve, reject) => {
          const timer = setTimeout(
            () => reject(new Error(`Worker did not report ${field}`)),
            10000,
          );
          worker.addEventListener('message', function receive({ data }) {
            if (data[field] !== undefined) {
              clearTimeout(timer);
              worker.removeEventListener('message', receive);
              resolve(data[field]);
            }
          });
        });
      const armed = wait('formulaCopyArmed');
      worker.postMessage({ method: 'cancelRequest', args: [-1], armFormulaCopy: true });
      await armed;
      const stalled = wait('formulaCopyStalled');
      const controller = new AbortController();
      const read = book
        .readCells(
          sheet.id,
          { startRow: 0, endRow: 1, startColumn: 1, endColumn: 2 },
          { signal: controller.signal },
        )
        .then(
          () => 'completed',
          (error: { code: string }) => error.code,
        );
      await stalled;
      const cancelled = wait('formulaCopyCancelled');
      controller.abort();
      await cancelled;
      worker.postMessage({ method: 'cancelRequest', args: [-1], releaseFormulaCopy: true });
      const code = await read;
      const count = wait('formulaCopyTimers');
      worker.postMessage({ method: 'cancelRequest', args: [-1], reportFormulaCopy: true });
      const timers = await count;
      const value = await sheet.range('B1').getValues();
      const id = book.id;
      const revision = book.revision;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      sheet = book.getSheets()[0];
      return {
        code,
        timers,
        value,
        revision,
        reopenedRevision: book.revision,
        reopened: await sheet.range('B1').getValues(),
      };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.code).toBe('ABORTED');
  expect(result.timers).toBe(1);
  expect(result.value).toEqual([[2550]]);
  expect(result.reopened).toEqual([[2550]]);
  expect(result.reopenedRevision).toBe(result.revision);
});

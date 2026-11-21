import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('bounds and clears the snapshot formula cache during real Worker file reads', async ({
  page,
}) => {
  test.setTimeout(60_000);
  const modulePath = `/@fs${fileURLToPath(new URL('../../packages/core/src/index.ts', import.meta.url))}`;
  await page.route('**/file-reader-formula-cache.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { createWorkbookFile, encodeFileCells, FileReader } from ${JSON.stringify(modulePath)};
      try {
        const count = 35000;
        const file = createWorkbookFile({sheets:[{name:'Data',rows:count,columns:1}]});
        const sheet = file.sheets[0];
        for (let base = 0; base < count; base += 64) {
          const cells = Array.from({length: Math.min(64,count-base)}, (_,i) => ({
            rowId:base+i,columnId:0,input:{type:'formula',expression:String(base+i)+'+1'}
          }));
          sheet.blocks.push({row:base/64,column:0,data:encodeFileCells(cells,new Map())});
        }
        const book = new FileReader(file);
        let checked = 0;
        for (const batch of book.getSheetById(sheet.id).range('A1:A35000').batches({maxCells:1024})) {
          for (const row of batch.values) {
            if (row[0] !== ++checked) throw new Error('Unexpected value '+checked);
          }
        }
        const bytes = book.calculator.cacheBytes;
        const first = book.value(sheet.id,0,0);
        book.dispose();
        const cleared = book.calculator.cacheBytes;
        self.postMessage({checked,bytes,first,cleared});
      } catch (error) { self.postMessage({error:String(error)}); }`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<{ checked: number; bytes: number; first: number; cleared: number }>(
        (resolve, reject) => {
          const worker = new Worker('/file-reader-formula-cache.worker.js', { type: 'module' });
          worker.onmessage = ({ data }) => {
            worker.terminate();
            resolve(data);
          };
          worker.onerror = (event) => {
            worker.terminate();
            reject(new Error(event.message));
          };
        },
      ),
  );
  expect(result.checked).toBe(35_000);
  expect(result.first).toBe(1);
  expect(result.bytes).toBeGreaterThan(3 * 1024 * 1024);
  expect(result.bytes).toBeLessThanOrEqual(4 * 1024 * 1024);
  expect(result.cleared).toBe(0);
});

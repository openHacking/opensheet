import { expect, test } from '@playwright/test';

test('staged scalar copy reads each source cell once with a zero-byte cache', async ({ page }) => {
  await page.route('**/selective-copy-read-test.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
import { Engine } from '/@fs${process.cwd()}/packages/core/src/engine.ts';
import { createSnapshot } from '/@fs${process.cwd()}/packages/core/src/model.ts';
import { Database } from '/@fs${process.cwd()}/packages/core/src/storage.ts';
(async () => {
  const snapshot = createSnapshot({sheets:[{name:'Data', rows:1000, columns:2}]});
  const sheet = snapshot.sheets[0];
  for (let row=0; row<1000; row++) sheet.cells[row+':0'] = {rowId:row,columnId:0,input:{type:'number',value:row}};
  const db = await Database.open('selective-read-test-'+crypto.randomUUID());
  const engine = new Engine(db,snapshot.workbookId,0);
  await engine.open(snapshot);
  try {
    const read = engine.stagedCell.bind(engine);
    let reads = 0;
    engine.stagedCell = (...args) => { reads++; return read(...args); };
    await engine.execute([{type:'core.cells.copySpecial',payload:{sheetId:sheet.id,mode:'values',range:{startRow:0,endRow:1000,startColumn:0,endColumn:1},target:{startRow:0,endRow:1000,startColumn:1,endColumn:2}}}]);
    const copyReads = reads;
    const result = await engine.read(sheet.id,{startRow:0,endRow:1000,startColumn:1,endColumn:2},true);
    const verified = Object.values(result.calculated).every((value,index)=>value===index);
    self.postMessage({copyReads, verified, count:Object.keys(result.calculated).length});
  } finally { engine.close(); }
})().catch(error=>self.postMessage({error:String(error)}));`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/selective-copy-read-test.worker.js', { type: 'module' });
        worker.onmessage = ({ data }) => {
          worker.terminate();
          resolve(data);
        };
        worker.onerror = (event) => {
          worker.terminate();
          reject(new Error(event.message));
        };
      }),
  );
  expect(result).toEqual({ copyReads: 1000, verified: true, count: 1000 });
});

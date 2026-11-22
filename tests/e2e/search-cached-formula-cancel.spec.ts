import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const cacheBytes of [0, 1024 * 1024]) {
  test(`cancelled search stops cached formula traversal with cache ${cacheBytes}`, async ({
    page,
  }) => {
    const module = (name: string) =>
      `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
    await page.route('**/search-cached-formula.worker.js', (route) =>
      route.fulfill({
        contentType: 'text/javascript',
        body: `
import { Engine } from ${JSON.stringify(module('engine'))};
import { createSnapshot } from ${JSON.stringify(module('model'))};
import { Database } from ${JSON.stringify(module('storage'))};
let engine, resume, armed = false, yields = 0;
const originalTimeout = globalThis.setTimeout;
globalThis.setTimeout = (callback, delay, ...args) => {
  if (armed && delay === 0) {
    yields++;
    if (yields === 1) return originalTimeout(() => { resume = () => callback(...args); postMessage({kind:'checkpoint'}); }, 0);
  }
  return originalTimeout(callback, delay, ...args);
};
self.onmessage = ({data}) => { if (data === 'cancel') { engine.cancel(); resume?.(); } };
(async () => {
  const snapshot = createSnapshot({sheets:[{name:'Data', rows:10, columns:3}]});
  const sheet = snapshot.sheets[0];
  for (let row = 0; row < 10; row++) {
    sheet.cells[row+':0'] = {rowId:row,columnId:0,input:{type:'number',value:1}};
    sheet.cells[row+':2'] = {rowId:row,columnId:2,input:{type:'formula',expression:'A'+(row+1)+'*1'}};
  }
  const source = ${cacheBytes === 0 ? "'A1:A10'" : "'C1:C10'"};
  sheet.cells['0:1'] = {rowId:0,columnId:1,input:{type:'formula',expression:'SUM('+Array(255).fill(source).join(',')+')'}};
  const db = await Database.open('search-cached-cancel-'+crypto.randomUUID());
  engine = new Engine(db, snapshot.workbookId, ${cacheBytes});
  await engine.open(snapshot);
  engine.begin(); armed = true;
  let cancelled;
  try { await engine.search(sheet.id, 'needle', -1, true); cancelled = 'completed'; }
  catch (error) { cancelled = error.code; }
  armed = false;
  const observedYields = yields;
  engine.begin();
  const retry = await engine.search(sheet.id, 'needle', -1, true);
  const value = await engine.read(sheet.id,{startRow:0,endRow:1,startColumn:1,endColumn:2},true);
  await engine.clear(); engine.close();
  postMessage({kind:'result', cancelled, observedYields, retry, value:value.calculated['0:1']});
})().catch(error => postMessage({kind:'error',error:String(error),stack:error.stack}));`,
      }),
    );
    await page.goto('/');
    const result = await page.evaluate(
      () =>
        new Promise<any>((resolve, reject) => {
          const worker = new Worker('/search-cached-formula.worker.js', { type: 'module' });
          const timer = setTimeout(() => {
            worker.terminate();
            reject(new Error('Search did not reach checkpoint'));
          }, 15000);
          worker.onerror = (event) => {
            clearTimeout(timer);
            worker.terminate();
            reject(new Error(event.message));
          };
          worker.onmessage = ({ data }) => {
            if (data.kind === 'checkpoint') {
              worker.postMessage('cancel');
              return;
            }
            clearTimeout(timer);
            worker.terminate();
            resolve(data);
          };
        }),
    );
    expect(result).toEqual({
      kind: 'result',
      cancelled: 'ABORTED',
      observedYields: 1,
      retry: { count: 0, match: undefined },
      value: 2550,
    });
  });
}

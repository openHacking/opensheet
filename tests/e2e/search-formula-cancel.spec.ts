import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('zero-cache formula search can be cancelled during a page read and retried', async ({
  page,
}) => {
  const module = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/formula-search-cancel.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
        import { Engine } from ${JSON.stringify(module('engine'))};
        import { createSnapshot } from ${JSON.stringify(module('model'))};
        import { Database } from ${JSON.stringify(module('storage'))};
        let engine, resume;
        self.onmessage = ({data}) => { if (data === 'cancel') { engine.cancel(); resume?.(); } };
        (async () => {
          const snapshot = createSnapshot({sheets:[{name:'Data', rows:3, columns:1}]});
          const sheet = snapshot.sheets[0];
          sheet.cells['0:0'] = {rowId:0, columnId:0, input:{type:'formula',expression:'A2+1'}};
          sheet.cells['1:0'] = {rowId:1, columnId:0, input:{type:'number',value:1}};
          const db = await Database.open('formula-search-cancel-' + crypto.randomUUID());
          engine = new Engine(db, snapshot.workbookId, 0);
          await engine.open(snapshot);
          const read = engine.stagedCell.bind(engine);
          let armed = true;
          engine.stagedCell = async (...args) => {
            if (armed) {
              armed = false;
              await new Promise(resolve => { resume = resolve; postMessage({kind:'checkpoint'}); });
            }
            return read(...args);
          };
          engine.begin();
          let cancelled;
          try { await engine.search(sheet.id, '2', -1, true); cancelled = 'completed'; }
          catch (error) { cancelled = error.code; }
          engine.begin();
          const retry = await engine.search(sheet.id, '2');
          await engine.clear(); engine.close();
          postMessage({kind:'result',cancelled,retry});
        })().catch(error => postMessage({kind:'error',error:String(error),stack:error.stack}));
      `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/formula-search-cancel.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Formula search did not reach its read checkpoint'));
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
    retry: { count: 1, match: { row: 0, column: 0 } },
  });
});

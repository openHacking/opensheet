import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('validation staging cancels a pending IndexedDB write and its ledger in a Worker', async ({
  page,
}) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/validation-stage-idb-abort.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
      import { encodeWorkbookFile } from ${JSON.stringify(modulePath('file-codec'))};
      import { Database } from ${JSON.stringify(modulePath('storage'))};
      import { stageValidationFile, upgradeValidationFile } from ${JSON.stringify(modulePath('validation-file-migration'))};
      (async () => {
        const snapshot = createSnapshot({sheets:[{name:'Data',rows:2,columns:2}]});
        snapshot.sheets[0].cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:5}};
        const file = upgradeValidationFile(encodeWorkbookFile(snapshot));
        file.sheets[0].validation.push({id:'positive',range:{startRow:0,endRow:1,startColumn:0,endColumn:1},rule:{kind:'number',operator:'gte',first:1,allowBlank:false,policy:'reject'}});
        const db = await Database.open('validation-stage-abort-' + crypto.randomUUID());
        try {
          await db.write('records', [['baseline', {value:7}]]);
          const abort = new AbortController();
          let cancelled = '', writes = 0;
          try {
            await stageValidationFile(file, async (_sheet, block, signal) => {
              writes++;
              const pending = db.write('records', [['staged', block]], [], signal);
              abort.abort();
              await pending;
            }, abort.signal);
          } catch (error) { cancelled = error.name; }
          const staged = await db.get('records', 'staged');
          const ledger = await db.get('sizes', 'staged');
          const baseline = await db.get('records', 'baseline');
          const afterRequest = new AbortController();
          const originalPut = IDBObjectStore.prototype.put;
          let recordSucceeded = false, ledgerSucceeded = false, cancelledAfterRequest = '';
          IDBObjectStore.prototype.put = function(...args) {
            const request = originalPut.apply(this, args);
            if (args[1] === 'staged-after-request') {
              if (this.name === 'records') request.addEventListener('success', () => { recordSucceeded = true; }, {once:true});
              if (this.name === 'sizes') request.addEventListener('success', () => { ledgerSucceeded = true; afterRequest.abort(); }, {once:true});
            }
            return request;
          };
          try {
            await stageValidationFile(file, async (_sheet, block, signal) => {
              await db.write('records', [['staged-after-request', block]], [], signal);
            }, afterRequest.signal);
          } catch (error) { cancelledAfterRequest = error.name; }
          finally { IDBObjectStore.prototype.put = originalPut; }
          const afterRequestRecord = await db.get('records', 'staged-after-request');
          const afterRequestLedger = await db.get('sizes', 'staged-after-request');
          await db.write('records', [['staged-after-request', {value:11}]]);
          const retried = await db.get('records', 'staged-after-request');
          const retrySizes = await db.recordSizes(['staged-after-request']);
          const after = new AbortController();
          await db.write('records', [['committed', {value:9}]], [], after.signal);
          after.abort();
          const committed = await db.get('records', 'committed');
          self.postMessage({cancelled,writes,stagedAbsent:staged === undefined,ledgerAbsent:ledger === undefined,
            recordSucceeded,ledgerSucceeded,cancelledAfterRequest,afterRequestRecordAbsent:afterRequestRecord === undefined,
            afterRequestLedgerAbsent:afterRequestLedger === undefined,baseline,committed,retried,retrySize:retrySizes[0]});
        } finally { db.close(); }
      })().catch(error => self.postMessage({error:String(error),stack:error.stack}));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-stage-idb-abort.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker timeout'));
        }, 10000);
        worker.onerror = (error) => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error(error.message));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timer);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result.error).toBeUndefined();
  expect(result).toMatchObject({
    cancelled: 'AbortError',
    writes: 1,
    stagedAbsent: true,
    ledgerAbsent: true,
    recordSucceeded: true,
    ledgerSucceeded: true,
    cancelledAfterRequest: 'AbortError',
    afterRequestRecordAbsent: true,
    afterRequestLedgerAbsent: true,
    baseline: { value: 7 },
    committed: { value: 9 },
    retried: { value: 11 },
  });
  expect(result.retrySize).toBeGreaterThan(0);
});

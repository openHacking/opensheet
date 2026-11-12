import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('stages conditional metadata and aborts a pending IndexedDB write in a Worker', async ({
  page,
}) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/conditional-stage-idb.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
      import { encodeWorkbookFile, validateWorkbookFile } from ${JSON.stringify(modulePath('file-codec'))};
      import { Database } from ${JSON.stringify(modulePath('storage'))};
      import { stageConditionalFile, stageConditionalDowngrade, upgradeConditionalFile, upgradeValidationFile } from ${JSON.stringify(modulePath('validation-file-migration'))};
      (async () => {
        const snapshot = createSnapshot({sheets:[{name:'Data',rows:2,columns:2}]});
        snapshot.sheets[0].cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:5}};
        const file = upgradeConditionalFile(upgradeValidationFile(encodeWorkbookFile(snapshot)));
        file.sheets[0].validation.push({id:'positive',range:{startRow:0,endRow:1,startColumn:0,endColumn:1},rule:{kind:'number',operator:'gte',first:1,allowBlank:false,policy:'reject'}});
        file.sheets[0].conditionalFormats.push({id:'highlight',range:{startRow:0,endRow:1,startColumn:0,endColumn:1},condition:{kind:'formula',expression:'A1>0'},style:{bold:true},stopIfTrue:true});
        const db = await Database.open('conditional-stage-' + crypto.randomUUID());
        try {
          await db.write('records', [['baseline', {value:7}]]);
          const abort = new AbortController();
          let cancelled = '', writes = 0, downgrade = '', downgradeWrites = 0, legacy = '';
          try { await stageConditionalFile(file, async (_sheet, block, signal) => {
            writes++;
            const pending = db.write('records', [['cancelled', block]], [], signal);
            abort.abort();
            await pending;
          }, abort.signal); } catch (error) { cancelled = error.name; }
          const record = await db.get('records', 'cancelled');
          const ledger = await db.get('sizes', 'cancelled');
          const baseline = await db.get('records', 'baseline');
          const metadata = await stageConditionalFile(file, async (_sheet, block) => { await db.write('records', [['staged', block]]); });
          const staged = await db.get('records', 'staged');
          const size = (await db.recordSizes(['staged']))[0];
          try { await stageConditionalDowngrade(file, async () => { downgradeWrites++; }); } catch (error) { downgrade = error.code; }
          try { validateWorkbookFile(file); } catch (error) { legacy = error.code; }
          self.postMessage({cancelled,writes,recordAbsent:record === undefined,ledgerAbsent:ledger === undefined,baseline,metadata,staged:!!staged,size,downgrade,downgradeWrites,legacy});
        } finally { db.close(); }
      })().catch(error => self.postMessage({error:String(error),stack:error.stack}));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/conditional-stage-idb.worker.js', { type: 'module' });
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
    recordAbsent: true,
    ledgerAbsent: true,
    baseline: { value: 7 },
    staged: true,
    downgrade: 'UNSUPPORTED_FEATURE',
    downgradeWrites: 0,
    legacy: 'UNSUPPORTED_VERSION',
  });
  expect(result.size).toBeGreaterThan(0);
  expect(result.metadata.schemaVersion).toBe(5);
  expect(result.metadata.sheets[0].validation).toHaveLength(1);
  expect(result.metadata.sheets[0].conditionalFormats).toEqual([
    {
      id: 'highlight',
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      condition: { kind: 'formula', expression: 'A1>0' },
      style: { bold: true },
      stopIfTrue: true,
    },
  ]);
  expect(result.metadata.sheets[0]).not.toHaveProperty('blocks');
});

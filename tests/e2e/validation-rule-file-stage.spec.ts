import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('rule-bearing validation file stages in a Worker and rejects invalid rules before output', async ({
  page,
}) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/validation-rule-file-stage.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
      import { encodeWorkbookFile } from ${JSON.stringify(modulePath('file-codec'))};
      import { stageValidationFile, upgradeValidationFile } from ${JSON.stringify(modulePath('validation-file-migration'))};
      (async () => {
        const snapshot = createSnapshot({sheets:[{name:'Data',rows:2,columns:2}]});
        snapshot.sheets[0].cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:5}};
        const file = upgradeValidationFile(encodeWorkbookFile(snapshot));
        file.sheets[0].validation.push({id:'positive',range:{startRow:0,endRow:1,startColumn:0,endColumn:1},rule:{kind:'formula',expression:'A1>0',allowBlank:false,policy:'reject'}});
        const pages = [];
        const metadata = await stageValidationFile(file, async (sheet, page) => pages.push({sheet,page}));
        file.sheets[0].validation[0].rule.expression = 'A1+';
        let invalid = '', writes = 0;
        try { await stageValidationFile(file, async () => { writes++; }); }
        catch (error) { invalid = error.code; }
        file.sheets[0].validation[0].rule.expression = 'A1>0';
        const abort = new AbortController();
        let stopped = false, cancelled = '';
        try {
          await stageValidationFile(file, async (_sheet, _page, signal) => new Promise((_resolve, reject) => {
            signal.addEventListener('abort', () => { stopped = true; reject(signal.reason); }, {once:true});
            setTimeout(() => abort.abort(), 0);
          }), abort.signal);
        } catch (error) { cancelled = error.name; }
        self.postMessage({pages:pages.length,rule:metadata.sheets[0].validation[0],invalid,writes,stopped,cancelled});
      })().catch(error => self.postMessage({error:String(error)}));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-rule-file-stage.worker.js', { type: 'module' });
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
    pages: 1,
    rule: { id: 'positive', rule: { expression: 'A1>0' } },
    invalid: 'INVALID_ARGUMENT',
    writes: 0,
    stopped: true,
    cancelled: 'AbortError',
  });
});

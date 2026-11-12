import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('internal named-range index and transforms remain bounded and isolated in Workers', async ({
  page,
}) => {
  test.setTimeout(120_000);
  const path = `/@fs${fileURLToPath(new URL('../../packages/core/src/named-range-model.ts', import.meta.url))}`;
  await page.route('**/named-range-model.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
    import { indexNamedRanges, transformNamedRanges, transformStagedNamedRanges, removeNamedRangeSheet } from ${JSON.stringify(path)};
    try {
      const sheets = [{id:'first',rows:10,columns:3},{id:'second',rows:10,columns:3}];
      const names = [];
      for (let i=0;i<4096;i++) {
        names.push({id:'g_'+i,name:'Range_'+i,scopeSheetId:null,target:{sheetId:'first',range:{startRow:2,startColumn:0,endRow:4,endColumn:1}}});
        names.push({id:'l_'+i,name:'range_'+i,scopeSheetId:'second',target:{sheetId:'second',range:{startRow:1,startColumn:0,endRow:2,endColumn:1}}});
      }
      const index = indexNamedRanges(names,sheets);
      for (let i=0;i<4096;i++) {
        if (index.resolve('RANGE_'+i,'first')?.id !== 'g_'+i || index.resolve('RANGE_'+i,'second')?.id !== 'l_'+i) throw new Error('scope '+i);
      }
      const changed = index.resolve('range_0','first'); changed.target.range.startRow=9;
      if (index.resolve('range_0','first').target.range.startRow !== 2 || names[0].target.range.startRow !== 2) throw new Error('isolation');
      const shifted = transformNamedRanges(names,sheets,'first','row','insert',1,1);
      for (let i=0;i<4096;i++) if (shifted[i*2].target.range.startRow !== 3 || shifted[i*2+1].target.range.startRow !== 1) throw new Error('shift '+i);
      const controller = new AbortController();
      // Force the elapsed-time checkpoint so cancellation is deterministic even on fast machines.
      let clock = 0;
      Object.defineProperty(performance, 'now', {configurable:true, value:() => clock += 10});
      const pending = transformStagedNamedRanges(names,sheets,'first','row','insert',1,1,controller.signal);
      setTimeout(() => controller.abort(), 0);
      let cancelled = false;
      try { await pending; } catch (error) { cancelled = error.name === 'AbortError'; }
      delete performance.now;
      if (!cancelled || names[0].target.range.startRow !== 2) throw new Error('cancellation');
      const staged = await transformStagedNamedRanges(names,sheets,'first','row','insert',1,1);
      if (JSON.stringify(staged) !== JSON.stringify(shifted)) throw new Error('staged mismatch');
      const removed = removeNamedRangeSheet(names,sheets,'first');
      if (removed.length !== 8192) throw new Error('count');
      for (let i=0;i<4096;i++) if (removed[i*2].target !== null || removed[i*2+1].target.sheetId !== 'second') throw new Error('remove '+i);
      let rejected=false; try { indexNamedRanges([...names,{...names[0]}],sheets); } catch { rejected=true; }
      self.postMessage({verified:8192,rejected,cancelled,original:names[0].target.range.startRow});
    } catch(error) { self.postMessage({error:String(error),stack:error.stack}); }
  `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/named-range-model.worker.js', { type: 'module' });
        const timeout = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker timeout'));
        }, 30000);
        worker.onerror = (error) => {
          clearTimeout(timeout);
          worker.terminate();
          reject(new Error(error.message));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timeout);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result).toEqual({ verified: 8192, rejected: true, cancelled: true, original: 2 });
});

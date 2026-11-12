import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('long text search yields and cancels before finishing a single page', async ({ page }) => {
  const module = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/long-text-search.worker.js*', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
        import { Engine } from ${JSON.stringify(module('engine'))};
        import { createSnapshot } from ${JSON.stringify(module('model'))};
        import { Database } from ${JSON.stringify(module('storage'))};
        let engine, resume;
        self.onmessage = ({data}) => { if(data === 'cancel') { engine.cancel(); resume?.(); } };
        (async()=>{
          const snapshot=createSnapshot({sheets:[{name:'Data',rows:64,columns:2}]});
          const sheet=snapshot.sheets[0];
          for(let row=0;row<64;row++) for(let column=0;column<2;column++) {
            const value='x'.repeat(32760)+(row===63&&column===1?'needle':'');
            sheet.cells[row+':'+column]={rowId:row,columnId:column,input:{type:'string',value}};
          }
          const db=await Database.open('long-search-'+crypto.randomUUID());
          engine=new Engine(db,snapshot.workbookId,0);
          await engine.open(snapshot);
          let visited=0;
          const block=engine.block.bind(engine);
          engine.block=async(...args)=>{
            const page=await block(...args), records=page.records.bind(page);
            page.records=function*(){ for(const cell of records()){ visited++; yield cell; } };
            return page;
          };
          const timer=self.setTimeout.bind(self);
          let armed=true;
          self.setTimeout=(callback,delay,...args)=>{
            if(armed&&delay===0){ armed=false; resume=()=>callback(...args); self.postMessage({kind:'checkpoint',visited}); return 0; }
            return timer(callback,delay,...args);
          };
          engine.begin();
          let cancelled;
          try { await engine.search(sheet.id,'needle',-1,true); cancelled='completed'; }
          catch(error) { cancelled=error.code; }
          engine.begin();
          const retry=await engine.search(sheet.id,'needle');
          const reverse=await engine.search(sheet.id,'needle',-1,false,'previous');
          await engine.clear(); engine.close();
          self.postMessage({kind:'result',cancelled,retry,reverse});
        })().catch(error=>self.postMessage({kind:'error',error:String(error),stack:error.stack}));
      `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/long-text-search.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Search did not yield'));
        }, 15000);
        let visited: number | undefined;
        worker.onerror = (event) => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error(event.message));
        };
        worker.onmessage = ({ data }) => {
          if (data.kind === 'checkpoint') {
            visited = data.visited;
            worker.postMessage('cancel');
            return;
          }
          clearTimeout(timer);
          worker.terminate();
          resolve({ ...data, visited });
        };
      }),
  );
  expect(result.kind).toBe('result');
  expect(result.visited).toBeGreaterThan(0);
  expect(result.visited).toBeLessThan(128);
  expect(result.cancelled).toBe('ABORTED');
  expect(result.retry).toEqual({ count: 1, match: { row: 63, column: 1 } });
  expect(result.reverse).toEqual(result.retry);
});

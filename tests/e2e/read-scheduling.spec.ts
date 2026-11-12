import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('range reads yield by visited cells and honor cancellation at the last checkpoint', async ({
  page,
}) => {
  const core = fileURLToPath(new URL('../../packages/core/src/', import.meta.url));
  await page.route('**/read-scheduling-test.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
import { Engine } from ${JSON.stringify(`/@fs${core}engine.ts`)};
import { createSnapshot } from ${JSON.stringify(`/@fs${core}model.ts`)};
import { Database } from ${JSON.stringify(`/@fs${core}storage.ts`)};
(async () => {
 const snapshot=createSnapshot({sheets:[{name:'Data',rows:1024,columns:2}]});
 const db=await Database.open('read-scheduling-'+crypto.randomUUID());
 const engine=new Engine(db,snapshot.workbookId);await engine.open(snapshot);
 const original=self.setTimeout;let timers=0, cancelAtYield=false;
 self.setTimeout=(fn,ms,...args)=>{timers++;return original(()=>{if(cancelAtYield)engine.cancel();fn(...args);},ms);};
 try {
  const id=snapshot.sheets[0].id;
  await engine.read(id,{startRow:0,endRow:1,startColumn:0,endColumn:1},false,false);
  const smallTimers=timers;timers=0;cancelAtYield=true;let error;
  try{await engine.read(id,{startRow:0,endRow:1024,startColumn:1,endColumn:2},true,false);}catch(e){error=e.code;}
  const cancelledTimers=timers;cancelAtYield=false;engine.begin();timers=0;
  const result=await engine.read(id,{startRow:0,endRow:1024,startColumn:1,endColumn:2},true,false);
  const values=Object.values(result.calculated);
  self.postMessage({smallTimers,cancelledTimers,error,retryTimers:timers,count:values.length,empty:values.every(v=>v===null)});
 }finally{self.setTimeout=original;engine.close();}
})().catch(error=>self.postMessage({error:String(error),stack:error.stack}));`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise((resolve, reject) => {
        const worker = new Worker('/read-scheduling-test.worker.js', { type: 'module' });
        worker.onerror = (event) => {
          worker.terminate();
          reject(new Error(event.message));
        };
        worker.onmessage = ({ data }) => {
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result).toEqual({
    smallTimers: 0,
    cancelledTimers: 1,
    error: 'ABORTED',
    retryTimers: 1,
    count: 1024,
    empty: true,
  });
});

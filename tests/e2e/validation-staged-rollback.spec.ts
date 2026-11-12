import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('injected staged validation rejection and cancellation preserve real IndexedDB state', async ({
  page,
}) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/validation-staged-rollback.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
        import { Engine } from ${JSON.stringify(modulePath('engine'))};
        import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
        import { Database } from ${JSON.stringify(modulePath('storage'))};
        const unhandled=[];
        self.addEventListener('unhandledrejection', event => {unhandled.push(String(event.reason));event.preventDefault()});
        const get = req => new Promise((resolve,reject) => {req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)});
        async function inspect(name,id) {
          const db = await get(indexedDB.open(name));
          try {
            const tx = db.transaction(['heads','records','sizes','scratch']);
            const range = IDBKeyRange.bound(id+'/',id+'/\uffff');
            const [head,records,sizes,scratch] = await Promise.all([
              get(tx.objectStore('heads').get(id)),
              get(tx.objectStore('records').getAllKeys(range)),
              get(tx.objectStore('sizes').getAllKeys(range)),
              get(tx.objectStore('scratch').getAllKeys(range))
            ]);
            return {head:Array.from(head),records,sizes,scratch};
          } finally { db.close(); }
        }
        async function run(mode) {
          const snapshot = createSnapshot({sheets:[{name:'Data',rows:2,columns:1}]});
          const sheet = snapshot.sheets[0];
          sheet.cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:5}};
          const name = 'validation-rollback-' + crypto.randomUUID();
          let db = await Database.open(name), engine = new Engine(db,snapshot.workbookId,0);
          await engine.open(snapshot);
          const before = engine.view(), rawBefore = await inspect(name,snapshot.workbookId);
          const originalPublish = engine.publish.bind(engine);
          const originalRead = engine.stagedCell.bind(engine);
          const controller = new AbortController();
          let release, entered, pending;
          const blocked = new Promise(resolve => {release=resolve});
          const started = new Promise(resolve => {entered=resolve});
          const range={startRow:0,endRow:1,startColumn:0,endColumn:1};
          const rules=[{id:'maximum',range,rule:{kind:'number',operator:'lte',first:10,allowBlank:false,policy:'reject'}}];
          let observed = 0;
          engine.publish = async (next,metadata) => {
            if(mode==='cancel') engine.stagedCell = (...args) => {
              pending = originalRead(...args).then(async () => {entered();await blocked;throw new Error('Late staged read failure')});
              return pending;
            };
            const checking = engine.validateStagedCells(next,metadata,sheet.id,rules,[{row:0,column:0}],controller.signal);
            if(mode==='cancel'){await started;controller.abort();}
            const report = await checking;
            observed = report.rejected;
            throw new Error('Injected validation rejection');
          };
          const command={type:'core.cells.set',payload:{sheetId:sheet.id,cells:[{row:0,column:0,input:{type:'number',value:12}}]}};
          let failure;
          try { await engine.execute([command]); } catch(error) { failure={name:error.name,message:error.message}; }
          const after=engine.view(), rawAfter=await inspect(name,snapshot.workbookId);
          release();
          let lateReadError;
          if(pending) try {await pending} catch(error) {lateReadError=error.message}
          await new Promise(resolve=>setTimeout(resolve,0));
          engine.stagedCell=originalRead; engine.publish=originalPublish;
          const valueAfter=(await engine.cell(sheet.id,0,0)).input.value;
          engine.close();
          db=await Database.open(name); engine=new Engine(db,snapshot.workbookId,0); await engine.open();
          const reopened=engine.view(), valueReopened=(await engine.cell(sheet.id,0,0)).input.value;
          await engine.execute([command]);
          const retry=(await engine.cell(sheet.id,0,0)).input.value;
          engine.close();
          return {mode, observed, failure, before:{revision:before.revision,bytes:before.bytes,canUndo:before.canUndo},
            after:{revision:after.revision,bytes:after.bytes,canUndo:after.canUndo},
            reopened:{revision:reopened.revision,bytes:reopened.bytes,canUndo:reopened.canUndo},
            rawEqual:JSON.stringify(rawBefore)===JSON.stringify(rawAfter),valueAfter,valueReopened,retry,lateReadError,unhandled:[...unhandled]};
        }
        (async()=>self.postMessage({results:[await run('reject'),await run('cancel')]}))()
          .catch(error=>self.postMessage({error:String(error),stack:error.stack}));
      `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-staged-rollback.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker did not finish'));
        }, 20000);
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
  for (const entry of result.results) {
    expect(entry.after).toEqual(entry.before);
    expect(entry.reopened).toEqual(entry.before);
    expect(entry.rawEqual).toBe(true);
    expect(entry.valueAfter).toBe(5);
    expect(entry.valueReopened).toBe(5);
    expect(entry.retry).toBe(12);
    expect(entry.unhandled).toEqual([]);
  }
  expect(result.results[0]).toMatchObject({
    mode: 'reject',
    observed: 1,
    failure: { message: 'Injected validation rejection' },
  });
  expect(result.results[1]).toMatchObject({
    mode: 'cancel',
    failure: { name: 'AbortError' },
    lateReadError: 'Late staged read failure',
  });
});

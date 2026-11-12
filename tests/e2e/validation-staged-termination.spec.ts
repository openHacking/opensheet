import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const stall of ['callback', 'idb'] as const) {
  test(`terminating a Worker during staged validation ${stall} recovers the committed head`, async ({
    page,
  }) => {
    const path = (name: string) =>
      `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
    await page.route('**/validation-staged-termination.worker.js*', (route) =>
      route.fulfill({
        contentType: 'text/javascript',
        body: `
        import { Engine } from ${JSON.stringify(path('engine'))};
        import { createSnapshot } from ${JSON.stringify(path('model'))};
        import { Database } from ${JSON.stringify(path('storage'))};
        const params=new URL(self.location.href).searchParams;
        const request=req=>new Promise((resolve,reject)=>{req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)});
        async function inspect(db,id) {
          const tx=db.db.transaction(['heads','records','sizes','scratch']);
          const range=IDBKeyRange.bound(id+'/',id+'/\uffff');
          const [head,records,sizes,scratch]=await Promise.all([
            request(tx.objectStore('heads').get(id)),request(tx.objectStore('records').getAllKeys(range)),
            request(tx.objectStore('sizes').getAllKeys(range)),request(tx.objectStore('scratch').getAllKeys(range))
          ]);
          return {head:Array.from(head),records,sizes,scratch};
        }
        (async()=>{
          const database=params.get('database');
          if(params.get('phase')==='reopen'){
            const id=params.get('id'), db=await Database.open(database), engine=new Engine(db,id,0);
            const view=await engine.open();
            const sheetId=view.snapshot.sheets[0].id;
            const value=(await engine.cell(sheetId,0,0)).input.value;
            const raw=await inspect(db,id);
            await engine.execute([{type:'core.cells.set',payload:{sheetId,cells:[{row:0,column:0,input:{type:'number',value:12}}]}}]);
            const retry=(await engine.cell(sheetId,0,0)).input.value;
            await engine.history(false);
            const undone=(await engine.cell(sheetId,0,0)).input.value;
            const result={kind:'reopened',revision:view.revision,bytes:view.bytes,canUndo:view.canUndo,value,raw,retry,undone};
            engine.close();self.postMessage(result);return;
          }
          const snapshot=createSnapshot({sheets:[{name:'Data',rows:2,columns:1}]});
          const sheet=snapshot.sheets[0], id=snapshot.workbookId;
          sheet.cells['0:0']={rowId:0,columnId:0,input:{type:'number',value:5}};
          const db=await Database.open(database), engine=new Engine(db,id,0);
          await engine.open(snapshot);
          const view=engine.view(), before=await inspect(db,id);
          const range={startRow:0,endRow:1,startColumn:0,endColumn:1};
          const rules=[{id:'maximum',range,rule:{kind:'number',operator:'lte',first:10,allowBlank:false,policy:'reject'}}];
          engine.publish=async(next,metadata)=>{
            const staged=await inspect(db,id);
            const notify=()=>{
              self.postMessage({kind:'staged',database,id,before,staged,
                view:{revision:view.revision,bytes:view.bytes,canUndo:view.canUndo},
                stall:params.get('stall')});
            };
            if(params.get('stall')==='idb'){
              const originalGet=db.get.bind(db), stagedRecord=next.blocks[sheet.id+'/0/0'].id;
              db.get=(store,key)=>{
                if(store!=='records'||key!==stagedRecord) return originalGet(store,key);
                return new Promise((resolve,reject)=>{
                  const tx=db.db.transaction('records','readonly'), records=tx.objectStore('records');
                  let notified=false;
                  const pulse=()=>{
                    const request=records.get(key);
                    request.onerror=()=>reject(request.error);
                    request.onsuccess=()=>{
                      if(!request.result){ reject(new Error('Staged record missing')); return; }
                      pulse(); // Keep another real request pending until Worker termination.
                      if(!notified){ notified=true; notify(); }
                    };
                  };
                  pulse();
                });
              };
            } else {
              engine.stagedCell=async()=>{ notify(); return new Promise(()=>{}); };
            }
            await engine.validateStagedCells(next,metadata,sheet.id,rules,[{row:0,column:0}]);
            throw new Error('Validation unexpectedly completed');
          };
          await engine.execute([{type:'core.cells.set',payload:{sheetId:sheet.id,cells:[{row:0,column:0,input:{type:'number',value:12}}]}}]);
        })().catch(error=>self.postMessage({kind:'error',error:String(error),stack:error.stack}));
      `,
      }),
    );
    await page.goto('/');
    const result = await page.evaluate(async (stall) => {
      const run = (url: string) =>
        new Promise<any>((resolve, reject) => {
          const worker = new Worker(url, { type: 'module' });
          const timer = setTimeout(() => {
            worker.terminate();
            reject(new Error('Worker timed out'));
          }, 15000);
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
        });
      const database = `validation-termination-${crypto.randomUUID()}`;
      const staged = await run(
        `/validation-staged-termination.worker.js?database=${database}&stall=${stall}`,
      );
      if (staged.kind !== 'staged') return { staged };
      const reopened = await run(
        `/validation-staged-termination.worker.js?phase=reopen&database=${database}&id=${staged.id}`,
      );
      return { staged, reopened };
    }, stall);
    expect(result.staged.kind).toBe('staged');
    expect(result.staged.stall).toBe(stall);
    expect(result.staged.staged.records.length).toBeGreaterThan(
      result.staged.before.records.length,
    );
    expect(result.staged.staged.sizes.length).toBeGreaterThan(result.staged.before.sizes.length);
    expect(result.staged.staged.head).toEqual(result.staged.before.head);
    expect(result.reopened.kind).toBe('reopened');
    expect(result.reopened.raw).toEqual(result.staged.before);
    expect(result.reopened).toMatchObject({ ...result.staged.view, value: 5 });
    expect(result.reopened).toMatchObject({ retry: 12, undone: 5 });
  });
}

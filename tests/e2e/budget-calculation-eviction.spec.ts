import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('budget formulas recalculate after their calculation-cache entries are evicted', async ({
  page,
}) => {
  const module = (path: string) => `/@fs${fileURLToPath(new URL(path, import.meta.url))}`;
  await page.route('**/budget-calculation-eviction.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
import { Engine } from ${JSON.stringify(module('../../packages/core/src/engine.ts'))};
import { Database } from ${JSON.stringify(module('../../packages/core/src/storage.ts'))};
import { decodeWorkbookFile } from ${JSON.stringify(module('../../packages/core/src/file-codec.ts'))};
import { createSnapshot } from ${JSON.stringify(module('../../packages/core/src/model.ts'))};
import budgetFile from ${JSON.stringify(module('../../apps/playground/src/budget-fixture.ts'))};
(async () => {
  const snapshot = decodeWorkbookFile(budgetFile);
  snapshot.workbookId='budget-calc-eviction-'+crypto.randomUUID();
  const [budget, notes] = snapshot.sheets;
  notes.cells['1:3']={rowId:1,columnId:3,input:{type:'formula',expression:"'Launch budget'!$D2*'Launch budget'!E$2"}};
  const filler=createSnapshot({sheets:[{name:'Filler',rows:512,columns:1}]}).sheets[0];
  for(let row=0;row<512;row++) filler.cells[row+':0']={rowId:row,columnId:0,input:{type:'formula',expression:row+'+1'}};
  snapshot.sheets.push(filler); snapshot.sheetOrder.push(filler.id);
  const database='budget-calc-eviction-'+crypto.randomUUID();
  let db=await Database.open(database), engine=new Engine(db,snapshot.workbookId,512*1024);
  await engine.open(snapshot);
  const rectangle=(row,column)=>({startRow:row,endRow:row+1,startColumn:column,endColumn:column+1});
  const value=async(sheet,row,column)=>(await engine.read(sheet,rectangle(row,column),true)).calculated[row+':'+column];
  const read=async()=>({budget:await value(budget.id,1,5),reference:await value(notes.id,1,3),total:await value(budget.id,10,5)});
  const id=notes.id+':1:3';
  const entries=()=>engine.calculator.cache.entries;
  const cache=engine.calculator.cache, get=cache.get.bind(cache);
  let targetHits=0;
  cache.get=(key)=>{const result=get(key); if(key===id && result!==undefined) targetHits++; return result;};
  const startRevision=engine.view().revision;
  const before=await read(), presentAfterFirst=entries().has(id), bytes=engine.calculator.cacheBytes;
  const hitsBeforeWarm=targetHits;
  const warmed=await read(), presentAfterWarm=entries().has(id), warmHit=targetHits>hitsBeforeWarm;
  const filled=await engine.read(filler.id,{startRow:0,endRow:512,startColumn:0,endColumn:1},true);
  const fillerCorrect=Array.from({length:512},(_,i)=>filled.calculated[i+':0']===i+1).every(Boolean);
  const presentAfterFill=entries().has(id), cacheBytes=engine.calculator.cacheBytes, cacheBudget=engine.calculator.cacheBudget;
  const afterEviction=await read(), presentAfterReread=entries().has(id);
  const revision=engine.view().revision;
  const command={type:'core.cells.set',payload:{sheetId:budget.id,cells:[{row:1,column:3,input:{type:'number',value:3}}]}};
  await engine.execute([command]); const edited=await read();
  const undo=await engine.history(false); const undone=await read();
  const redo=await engine.history(true); const redone=await read(), finalRevision=engine.view().revision;
  engine.close(); db=await Database.open(database); engine=new Engine(db,snapshot.workbookId,512*1024); await engine.open();
  const reopened=await read(), reopenedRevision=engine.view().revision;
  await engine.clear(); engine.close();
  postMessage({before,warmed,afterEviction,edited,undone,redone,reopened,presentAfterFirst,presentAfterWarm,
    presentAfterFill,presentAfterReread,warmHit,bytes,cacheBytes,cacheBudget,fillerCorrect,startRevision,revision,finalRevision,reopenedRevision,
    undo:!!undo.commit,redo:!!redo.commit});
})().catch(error=>postMessage({error:String(error),stack:error.stack}));`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/budget-calculation-eviction.worker.js', { type: 'module' });
        const timeout = setTimeout(() => {
          worker.terminate();
          reject(new Error('Calculation eviction test timed out'));
        }, 20000);
        worker.onerror = (event) => {
          clearTimeout(timeout);
          worker.terminate();
          reject(new Error(event.message));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timeout);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result.error).toBeUndefined();
  const initial = { budget: 2400, reference: 2400, total: 13097 };
  const edited = { budget: 7200, reference: 7200, total: 17897 };
  expect(result.before).toEqual(initial);
  expect(result.warmed).toEqual(initial);
  expect(result.afterEviction).toEqual(initial);
  expect(result.edited).toEqual(edited);
  expect(result.undone).toEqual(initial);
  expect(result.redone).toEqual(edited);
  expect(result.reopened).toEqual(edited);
  expect(result.presentAfterFirst).toBe(true);
  expect(result.presentAfterWarm).toBe(true);
  expect(result.warmHit).toBe(true);
  expect(result.presentAfterFill).toBe(false);
  expect(result.presentAfterReread).toBe(true);
  expect(result.bytes).toBeGreaterThan(0);
  expect(result.cacheBytes).toBeLessThanOrEqual(result.cacheBudget);
  expect(result.fillerCorrect).toBe(true);
  expect(result.undo).toBe(true);
  expect(result.redo).toBe(true);
  expect(result.revision).toBe(result.startRevision);
  expect(result.finalRevision).toBe(result.revision + 3);
  expect(result.reopenedRevision).toBe(result.finalRevision);
});

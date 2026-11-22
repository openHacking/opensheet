import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const operation of ['sort', 'filter', 'copy'] as const) {
  test(`cancelled ${operation} stops cached formula traversal before publication`, async ({
    page,
  }) => {
    const module = (name: string) =>
      `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
    await page.route('**/staged-cached-formula.worker.js', (route) =>
      route.fulfill({
        contentType: 'text/javascript',
        body: `
import { Engine } from ${JSON.stringify(module('engine'))};
import { createSnapshot } from ${JSON.stringify(module('model'))};
import { Database } from ${JSON.stringify(module('storage'))};
let engine, resume, armed = false, yields = 0;
const originalTimeout = globalThis.setTimeout;
globalThis.setTimeout = (callback, delay, ...args) => {
  if (armed && delay === 0) {
    yields++;
    if (yields === 1) return originalTimeout(() => { resume = () => callback(...args); postMessage({kind:'checkpoint'}); }, 0);
  }
  return originalTimeout(callback, delay, ...args);
};
self.onmessage = ({data}) => { if (data === 'cancel') { engine.cancel(); resume?.(); } };
(async () => {
  const snapshot = createSnapshot({sheets:[{name:'Data', rows:10, columns:4}]});
  const sheet = snapshot.sheets[0];
  for (let row = 0; row < 10; row++) {
    sheet.cells[row+':0'] = {rowId:row,columnId:0,input:{type:'number',value:1}};
    sheet.cells[row+':2'] = {rowId:row,columnId:2,input:{type:'formula',expression:'A'+(row+1)+'*1'}};
  }
  sheet.cells['0:1'] = {rowId:0,columnId:1,input:{type:'formula',expression:'SUM('+Array(255).fill('C1:C10').join(',')+')'}};
  sheet.cells['1:1'] = {rowId:1,columnId:1,input:{type:'number',value:0}};
  const database = 'staged-cached-cancel-'+crypto.randomUUID();
  let db = await Database.open(database);
  engine = new Engine(db, snapshot.workbookId, 1024*1024);
  await engine.open(snapshot);
  const range = {startRow:0,endRow:2,startColumn:1,endColumn:2};
  const operation = ${JSON.stringify(operation)};
  const command = operation === 'sort'
    ? {type:'core.sheet.sort',payload:{sheetId:sheet.id,range,column:1,direction:'asc'}}
    : operation === 'filter'
      ? {type:'core.sheet.filter',payload:{sheetId:sheet.id,filter:{mode:'all',headerRows:0,criteria:[{column:1,operator:'gt',value:0}]}}}
      : {type:'core.cells.copySpecial',payload:{sheetId:sheet.id,range:{...range,endRow:1},target:{startRow:0,endRow:1,startColumn:3,endColumn:4},mode:'values'}};
  const keys = async () => {
    const raw = db.db;
    const result = {};
    for (const name of ['records','sizes','scratch']) result[name] = await new Promise((resolve,reject) => {
      const request = raw.transaction(name).objectStore(name).getAllKeys();
      request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error);
    });
    return JSON.stringify(result);
  };
  const before = engine.view();
  const beforeHead = JSON.stringify(await db.get('heads', snapshot.workbookId));
  const beforeKeys = await keys();
  engine.begin(); armed = true;
  let cancelled;
  try { await engine.execute([command]); cancelled = 'completed'; }
  catch (error) { cancelled = error.code; }
  armed = false;
  const observedYields = yields;
  const after = engine.view();
  const rawSame = beforeHead === JSON.stringify(await db.get('heads', snapshot.workbookId));
  const keysSame = beforeKeys === await keys();
  const bytesSame = before.bytes === after.bytes;
  engine.begin();
  const old = await engine.read(sheet.id,{startRow:0,endRow:2,startColumn:1,endColumn:2},true);
  await engine.execute([command]);
  const applied = engine.view();
  const undo = await engine.history(false);
  const undone = engine.view();
  const redo = await engine.history(true);
  const redone = engine.view();
  const all = {startRow:0,endRow:10,startColumn:0,endColumn:4};
  const beforeReopen = await engine.read(sheet.id, all, true);
  engine.close();
  db = await Database.open(database);
  engine = new Engine(db, snapshot.workbookId, 1024*1024);
  await engine.open();
  const reopened = engine.view();
  const afterReopen = await engine.read(sheet.id, all, true);
  const reopenedSame = reopened.revision === redone.revision && JSON.stringify(reopened.filtered) === JSON.stringify(redone.filtered) && JSON.stringify(afterReopen.calculated) === JSON.stringify(beforeReopen.calculated);
  await engine.clear(); engine.close();
  postMessage({kind:'result',cancelled,observedYields,before:before.revision,after:after.revision,rawSame,keysSame,bytesSame,reopenedSame,old:[old.calculated['0:1'],old.calculated['1:1']],applied:applied.revision,undo:!!undo.commit,undone:undone.revision,redo:!!redo.commit,redone:redone.revision});
})().catch(error => postMessage({kind:'error',error:String(error),stack:error.stack}));`,
      }),
    );
    await page.goto('/');
    const result = await page.evaluate(
      () =>
        new Promise<any>((resolve, reject) => {
          const worker = new Worker('/staged-cached-formula.worker.js', { type: 'module' });
          const timer = setTimeout(() => {
            worker.terminate();
            reject(new Error('Staged operation did not reach checkpoint'));
          }, 20000);
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
    expect(result.kind).toBe('result');
    expect(result.cancelled).toBe('ABORTED');
    expect(result.observedYields).toBe(1);
    expect(result.after).toBe(result.before);
    expect(result.rawSame).toBe(true);
    expect(result.keysSame).toBe(true);
    expect(result.bytesSame).toBe(true);
    expect(result.reopenedSame).toBe(true);
    expect(result.old).toEqual([2550, 0]);
    expect(result.applied).toBe(result.before + 1);
    expect(result.undo).toBe(true);
    expect(result.redo).toBe(true);
    expect(result.undone).toBe(result.applied + 1);
    expect(result.redone).toBe(result.undone + 1);
  });
}

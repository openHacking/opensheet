import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('cancelled staged validation formula does not publish its candidate', async ({ page }) => {
  const module = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/validation-formula-cancel.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
import { Engine } from ${JSON.stringify(module('engine'))};
import { createSnapshot } from ${JSON.stringify(module('model'))};
import { Database } from ${JSON.stringify(module('storage'))};
let engine, resume, armed = false, yields = 0;
const timeout = globalThis.setTimeout;
globalThis.setTimeout = (callback, delay, ...args) => {
  if (armed && delay === 0 && ++yields === 1)
    return timeout(() => { resume = () => callback(...args); postMessage({kind:'checkpoint'}); }, 0);
  return timeout(callback, delay, ...args);
};
self.onmessage = ({data}) => { if (data === 'cancel') { engine.cancel(); resume(); } };
(async () => {
  const snapshot = createSnapshot({sheets:[{name:'Data',rows:10,columns:2}]});
  const sheet = snapshot.sheets[0];
  for (let row = 0; row < 10; row++)
    sheet.cells[row+':0'] = {rowId:row,columnId:0,input:{type:'formula',expression:'1'}};
  sheet.cells['0:1'] = {rowId:0,columnId:1,input:{type:'number',value:3}};
  const name = 'validation-formula-cancel-'+crypto.randomUUID();
  let db = await Database.open(name);
  engine = new Engine(db,snapshot.workbookId,1024*1024);
  await engine.open(snapshot);
  const keys = async () => {
    const result = {};
    for (const store of ['records','sizes','scratch']) result[store] = await new Promise((resolve,reject) => {
      const request = db.db.transaction(store).objectStore(store).getAllKeys();
      request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error);
    });
    return JSON.stringify(result);
  };
  const before = engine.view(), head = JSON.stringify(await db.get('heads',snapshot.workbookId)), beforeKeys = await keys();
  const publish = engine.publish.bind(engine);
  const range={startRow:0,endRow:1,startColumn:1,endColumn:2};
  const rules=[{id:'long',range,rule:{kind:'formula',expression:'SUM('+Array(255).fill('A1:A10').join(',')+')=2550',allowBlank:false,policy:'reject'}}];
  let report;
  engine.publish = async (next,metadata,ranges) => {
    armed = true;
    try { report = await engine.validateStagedCells(next,metadata,sheet.id,rules,[{row:0,column:1}]); }
    finally { armed = false; }
    if (report.rejected) throw new Error('Validation rejected');
    return publish(next,metadata,ranges);
  };
  const command={type:'core.cells.set',payload:{sheetId:sheet.id,cells:[{row:0,column:1,input:{type:'number',value:12}}]}};
  engine.begin();
  let failure;
  try { await engine.execute([command]); failure='success'; } catch(error) { failure=error.code; }
  armed = false;
  const observedYields=yields, after=engine.view();
  const headSame=head===JSON.stringify(await db.get('heads',snapshot.workbookId));
  const keysSame=beforeKeys===await keys();
  engine.begin();
  const cancelledValue=(await engine.cell(sheet.id,0,1)).input.value;
  await engine.execute([command]);
  const applied=engine.view(), appliedValue=(await engine.cell(sheet.id,0,1)).input.value;
  const undo=await engine.history(false), undoValue=(await engine.cell(sheet.id,0,1)).input.value;
  const redo=await engine.history(true), redoValue=(await engine.cell(sheet.id,0,1)).input.value;
  const final=engine.view();
  engine.close(); db=await Database.open(name); engine=new Engine(db,snapshot.workbookId,1024*1024); await engine.open();
  const reopened=engine.view(), reopenedValue=(await engine.cell(sheet.id,0,1)).input.value;
  await engine.clear(); engine.close();
  postMessage({kind:'result',failure,observedYields,headSame,keysSame,bytesSame:before.bytes===after.bytes,
    before:before.revision,after:after.revision,applied:applied.revision,final:final.revision,reopened:reopened.revision,
    cancelledValue,appliedValue,undo:!!undo.commit,undoValue,redo:!!redo.commit,redoValue,reopenedValue,report});
})().catch(error=>postMessage({kind:'error',error:String(error),stack:error.stack}));`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-formula-cancel.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Validation did not finish'));
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
  expect(result.failure).toBe('ABORTED');
  expect(result.observedYields).toBe(1);
  expect(result.headSame).toBe(true);
  expect(result.keysSame).toBe(true);
  expect(result.bytesSame).toBe(true);
  expect(result.after).toBe(result.before);
  expect(result.applied).toBe(result.before + 1);
  expect(result.reopened).toBe(result.final);
  expect(result).toMatchObject({
    cancelledValue: 3,
    appliedValue: 12,
    undo: true,
    undoValue: 3,
    redo: true,
    redoValue: 12,
    reopenedValue: 12,
    report: { rejected: 0, warnings: 0 },
  });
});

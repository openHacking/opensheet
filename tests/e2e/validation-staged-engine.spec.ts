import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('internal staged validation reads the unpublished transaction in a real Worker', async ({
  page,
}) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/validation-staged-engine.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
        import { Engine } from ${JSON.stringify(modulePath('engine'))};
        import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
        import { Database } from ${JSON.stringify(modulePath('storage'))};
        (async () => {
          const snapshot = createSnapshot({ sheets: [{name:'Data',rows:2,columns:1},{name:'Limits',rows:1,columns:1}] });
          const [data, limits] = snapshot.sheets;
          data.cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:5}};
          limits.cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:10}};
          const db = await Database.open('validation-staged-browser-' + crypto.randomUUID());
          const engine = new Engine(db, snapshot.workbookId, 0);
          await engine.open(snapshot);
          try {
            const range = {startRow:0,endRow:1,startColumn:0,endColumn:1};
            const rules = [
              {id:'minimum',range,rule:{kind:'number',operator:'gte',first:10,allowBlank:false,policy:'reject'}},
              {id:'limit',range,rule:{kind:'formula',expression:'A1<Limits!$A$1',allowBlank:false,policy:'warn'}}
            ];
            const positions = [{row:0,column:0}];
            const blank = await engine.validateStagedCells(engine.head, engine.metadata, data.id,
              [{id:'required',range:{startRow:1,endRow:2,startColumn:0,endColumn:1},
                rule:{kind:'formula',expression:'TRUE',allowBlank:false,policy:'reject'}}],
              [{row:1,column:0}]);
            const publish = engine.publish.bind(engine);
            let observation;
            engine.publish = async (next, metadata, ranges) => {
              const old = await engine.validateStagedCells(engine.head, engine.metadata, data.id, rules, positions);
              const staged = await engine.validateStagedCells(next, metadata, data.id, rules, positions);
              const fromRanges = await engine.validateStagedRanges(next, metadata, data.id, rules, [range, range]);
              observation = {old, staged, fromRanges, live: (await engine.cell(data.id, 0, 0)).input.value, revision: engine.view().revision};
              return publish(next, metadata, ranges);
            };
            await engine.execute([
              {type:'core.cells.set',payload:{sheetId:data.id,cells:[{row:0,column:0,input:{type:'number',value:12}}]}},
              {type:'core.cells.set',payload:{sheetId:limits.id,cells:[{row:0,column:0,input:{type:'number',value:8}}]}}
            ]);
            self.postMessage({observation, blank, committed:(await engine.cell(data.id,0,0)).input.value});
          } finally { engine.close(); }
        })().catch(error => self.postMessage({error:String(error), stack:error.stack}));
      `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-staged-engine.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker did not finish'));
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
  expect(result.blank).toMatchObject({
    rejected: 1,
    warnings: 0,
    details: [{ ruleId: 'required' }],
  });
  expect(result.observation).toMatchObject({
    old: { rejected: 1, warnings: 0, details: [{ ruleId: 'minimum' }] },
    staged: { rejected: 0, warnings: 1, details: [{ ruleId: 'limit' }] },
    fromRanges: { rejected: 0, warnings: 1, details: [{ ruleId: 'limit' }] },
    live: 5,
    revision: 0,
  });
  expect(result.committed).toBe(12);
});

test('engine validation can cancel during sheet-bound checks before a late invalid position', async ({
  page,
}) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/validation-engine-positions.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
        import { Engine } from ${JSON.stringify(modulePath('engine'))};
        import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
        import { Database } from ${JSON.stringify(modulePath('storage'))};
        (async () => {
          const snapshot = createSnapshot({sheets:[{name:'Data',rows:65536,columns:1}]});
          const sheet = snapshot.sheets[0];
          const engine = new Engine(await Database.open('validation-position-' + crypto.randomUUID()), snapshot.workbookId, 0);
          await engine.open(snapshot);
          try {
            const positions = Array.from({length:65536},(_,row)=>({row,column:0}));
            positions[positions.length - 1].row = 65536;
            const revision = engine.view().revision;
            let reads = 0;
            const read = engine.stagedCell.bind(engine);
            engine.stagedCell = (...args) => { reads++; return read(...args); };
            const controller = new AbortController();
            let clock = 0;
            Object.defineProperty(performance,'now',{configurable:true,value:()=>clock+=10});
            setTimeout(()=>controller.abort(),0);
            let cancelled;
            try { await engine.validateStagedCells(engine.head,engine.metadata,sheet.id,[],positions,controller.signal); cancelled='success'; }
            catch (error) { cancelled=error.name; }
            delete performance.now;
            let invalid;
            try { await engine.validateStagedCells(engine.head,engine.metadata,sheet.id,[],positions); invalid='success'; }
            catch (error) { invalid=error.code; }
            positions[positions.length - 1].row = 65535;
            const retry = await engine.validateStagedCells(engine.head,engine.metadata,sheet.id,[],positions);
            self.postMessage({cancelled,invalid,retry,reads,revision,after:engine.view().revision});
          } finally { engine.close(); }
        })().catch(error=>self.postMessage({error:String(error),stack:error.stack}));`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-engine-positions.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker timed out'));
        }, 30000);
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
  expect(result).toEqual({
    cancelled: 'AbortError',
    invalid: 'INVALID_ARGUMENT',
    retry: { rejected: 0, warnings: 0, details: [], truncated: false },
    reads: 0,
    revision: 0,
    after: 0,
  });
});

test('staged range enumeration can cancel a large disjoint scan and retry', async ({ page }) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/validation-range-cancel.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
        import { Engine } from ${JSON.stringify(modulePath('engine'))};
        import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
        import { Database } from ${JSON.stringify(modulePath('storage'))};
        (async () => {
          const snapshot = createSnapshot({sheets:[{name:'Data',rows:2,columns:1}]});
          const sheet = snapshot.sheets[0];
          const engine = new Engine(await Database.open('validation-range-'+crypto.randomUUID()),snapshot.workbookId,0);
          await engine.open(snapshot);
          try {
            const rule = {kind:'number',operator:'gte',first:1,allowBlank:false,policy:'reject'};
            const rules = Array.from({length:512},(_,i)=>({id:String(i),range:{startRow:0,endRow:1,startColumn:0,endColumn:1},rule}));
            const ranges = Array.from({length:1800},()=>({startRow:1,endRow:2,startColumn:0,endColumn:1}));
            let reads = 0;
            const read = engine.stagedCell.bind(engine);
            engine.stagedCell = (...args)=>{reads++;return read(...args);};
            const controller = new AbortController();
            let clock = 0;
            Object.defineProperty(performance,'now',{configurable:true,value:()=>clock+=10});
            setTimeout(()=>controller.abort(),0);
            let cancelled;
            try { await engine.validateStagedRanges(engine.head,engine.metadata,sheet.id,rules,ranges,controller.signal); cancelled='success'; }
            catch (error) { cancelled=error.name; }
            delete performance.now;
            const retry = await engine.validateStagedRanges(engine.head,engine.metadata,sheet.id,rules,ranges);
            self.postMessage({cancelled,retry,reads,revision:engine.view().revision});
          } finally {engine.close();}
        })().catch(error=>self.postMessage({error:String(error),stack:error.stack}));`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-range-cancel.worker.js', { type: 'module' });
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
    cancelled: 'AbortError',
    retry: { rejected: 0, warnings: 0, details: [], truncated: false },
    reads: 0,
    revision: 0,
  });
});

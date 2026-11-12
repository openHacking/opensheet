import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('resolves staged conditional appearance without publishing it in a real Worker', async ({
  page,
}) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/conditional-staged-engine.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { Engine } from ${JSON.stringify(modulePath('engine'))};
      import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
      import { Database } from ${JSON.stringify(modulePath('storage'))};
      (async () => {
        const snapshot = createSnapshot({sheets:[{name:'Data',rows:5,columns:1},{name:'Limits',rows:1,columns:1}]});
        const [data, limits] = snapshot.sheets;
        snapshot.styles.style1 = {color:'#111111',bold:true};
        data.cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:5}};
        data.cells['1:0'] = {rowId:1,columnId:0,input:{type:'formula',expression:'A1*2'},styleId:'style1'};
        data.cells['3:0'] = {rowId:3,columnId:0,input:{type:'error',code:'#DIV/0!'}};
        data.cells['4:0'] = {rowId:4,columnId:0,input:{type:'formula',expression:'A5'}};
        limits.cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:10}};
        const name = 'conditional-staged-engine-' + crypto.randomUUID();
        let db = await Database.open(name), engine = new Engine(db, snapshot.workbookId, 0);
        await engine.open(snapshot);
        try {
          const rules = [{id:'over',range:{startRow:0,endRow:2,startColumn:0,endColumn:1},condition:{kind:'formula',expression:'A1>Limits!$A$1'},style:{background:'#ff0000'},stopIfTrue:false}];
          const positions = [{row:0,column:0},{row:1,column:0}];
          const run = () => engine.conditionalStagedViewport(engine.head, engine.metadata, data.id, rules, positions);
          await engine.read(data.id, {startRow:0,endRow:2,startColumn:0,endColumn:1});
          const before = await run();
          const publish = engine.publish.bind(engine);
          let observed;
          engine.publish = async (next, metadata, ranges) => {
            observed = {old:await run(), staged:await engine.conditionalStagedViewport(next,metadata,data.id,rules,positions), revision:engine.view().revision, live:(await engine.cell(data.id,0,0)).input.value};
            return publish(next,metadata,ranges);
          };
          await engine.execute([
            {type:'core.cells.set',payload:{sheetId:data.id,cells:[{row:0,column:0,input:{type:'number',value:12}}]}},
            {type:'core.cells.set',payload:{sheetId:limits.id,cells:[{row:0,column:0,input:{type:'number',value:8}}]}}
          ]);
          engine.publish = publish;
          const committed = await run();
          await engine.history(false);
          const undone = await run();
          await engine.history(true);
          const redone = await run();
          engine.close();
          db = await Database.open(name);
          engine = new Engine(db, snapshot.workbookId, 0);
          await engine.open();
          const reopened = await run();
          const errorRules = [{id:'errors',range:{startRow:0,endRow:5,startColumn:0,endColumn:1},condition:{kind:'formula',expression:'IFERROR(A1,42)=42'},style:{italic:true},stopIfTrue:false}];
          const errorStyles = await engine.conditionalStagedViewport(engine.head,engine.metadata,data.id,errorRules,Array.from({length:5},(_,row)=>({row,column:0})));
          const numeric = [{id:'positive',range:{startRow:0,endRow:1,startColumn:0,endColumn:1},condition:{kind:'number',operator:'gt',first:0},style:{bold:true},stopIfTrue:false}];
          const repeated = Array.from({length:3},()=>({row:0,column:0}));
          const load = engine.block.bind(engine); let pageReads = 0;
          engine.block = async (...args) => { pageReads++; return load(...args); };
          const reused = await engine.conditionalStagedViewport(engine.head,engine.metadata,data.id,numeric,repeated);
          const reusedReads = pageReads; pageReads = 0;
          engine.block = async (...args) => { pageReads++; const page = await load(...args); Object.defineProperty(page,'bytes',{value:1024*1024+1}); return page; };
          const capped = await engine.conditionalStagedViewport(engine.head,engine.metadata,data.id,numeric,repeated);
          engine.block = load;
          self.postMessage({before,observed,committed,undone,redone,reopened,errorStyles,reused,reusedReads,capped,cappedReads:pageReads});
        } finally { engine.close(); }
      })().catch(error => self.postMessage({error:String(error),stack:error.stack}));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/conditional-staged-engine.worker.js', { type: 'module' });
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
  const oldStyles = [{}, { color: '#111111', bold: true }];
  const newStyles = [
    { background: '#ff0000' },
    { color: '#111111', bold: true, background: '#ff0000' },
  ];
  expect(result).toEqual({
    before: oldStyles,
    observed: { old: oldStyles, staged: newStyles, revision: 0, live: 5 },
    committed: newStyles,
    undone: oldStyles,
    redone: newStyles,
    reopened: newStyles,
    errorStyles: [{}, { color: '#111111', bold: true }, {}, { italic: true }, {}],
    reused: [{ bold: true }, { bold: true }, { bold: true }],
    reusedReads: 1,
    capped: [{ bold: true }, { bold: true }, { bold: true }],
    cappedReads: 3,
  });
});

test('does not retain a real oversized decoded page in a Worker', async ({ page }) => {
  const modulePath = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/conditional-oversized-page.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { Engine } from ${JSON.stringify(modulePath('engine'))};
      import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
      import { Database } from ${JSON.stringify(modulePath('storage'))};
      (async () => {
        const snapshot=createSnapshot({sheets:[{name:'Data',rows:40,columns:1}]});
        const sheet=snapshot.sheets[0];
        for(let row=0;row<40;row++) sheet.cells[row+':0']={rowId:row,columnId:0,input:{type:'string',value:row+':'+'x'.repeat(32760)}};
        const db=await Database.open('conditional-real-page-'+crypto.randomUUID());
        const engine=new Engine(db,snapshot.workbookId,0);
        await engine.open(snapshot);
        try {
          const original=engine.block.bind(engine), sizes=[];
          engine.block=async (...args)=>{const page=await original(...args);sizes.push(page.bytes);return page;};
          const rules=[{id:'text',range:{startRow:0,endRow:40,startColumn:0,endColumn:1},condition:{kind:'text',operator:'contains',value:'x',caseSensitive:true},style:{bold:true},stopIfTrue:false}];
          const styles=await engine.conditionalStagedViewport(engine.head,engine.metadata,sheet.id,rules,[0,1,2].map(row=>({row,column:0})));
          self.postMessage({sizes,styles});
        } finally {engine.close();}
      })().catch(error=>self.postMessage({error:String(error),stack:error.stack}));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/conditional-oversized-page.worker.js', { type: 'module' });
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
  expect(result.styles).toEqual([{ bold: true }, { bold: true }, { bold: true }]);
  expect(result.sizes).toHaveLength(3);
  expect(result.sizes.every((size: number) => size > 1024 * 1024)).toBe(true);
});

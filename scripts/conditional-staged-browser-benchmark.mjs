import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { arch, cpus, platform, totalmem } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2],
  engine = process.argv[3] ?? 'chromium';
const launchers = { chromium, firefox, webkit };
if (!output || !launchers[engine])
  throw new Error(
    'Usage: node scripts/conditional-staged-browser-benchmark.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const modulePath = (name) => `/@fs${resolve(`packages/core/src/${name}.ts`)}`;
const report = {
  startedAt: new Date().toISOString(),
  engine,
  source: sourceFingerprint(),
  host: { platform: platform(), arch: arch(), cpu: cpus()[0].model, memoryBytes: totalmem() },
  scope:
    'Development-server source in a real Worker and IndexedDB; 1,000 numeric cells, one relative-reference formula rule, three rounds each with zero and 64 MiB cache reservations. Fixture setup excluded; all styles checked after timing. No renderer, main-thread frame timing, peak process memory, external devices or durable rule publication.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const browser = await launchers[engine].launch();
try {
  report.browser = browser.version();
  const page = await browser.newPage();
  await page.route('**/conditional-staged-benchmark.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
    import { Engine } from ${JSON.stringify(modulePath('engine'))};
    import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
    import { Database } from ${JSON.stringify(modulePath('storage'))};
    (async () => {
      const results = [];
      for (const cache of [0, 64*1024*1024]) {
        const snapshot = createSnapshot({sheets:[{name:'Data',rows:1000,columns:1}]});
        const sheet = snapshot.sheets[0];
        for (let row=0; row<1000; row++) sheet.cells[row+':0'] = {rowId:row,columnId:0,input:{type:'number',value:row}};
        const db = await Database.open('conditional-staged-bench-'+crypto.randomUUID());
        const engine = new Engine(db,snapshot.workbookId,cache);
        await engine.open(snapshot);
        try {
          const rules = [{id:'positive',range:{startRow:0,endRow:1000,startColumn:0,endColumn:1},condition:{kind:'formula',expression:'A1>0'},style:{bold:true},stopIfTrue:false}];
          const positions = Array.from({length:1000},(_,row)=>({row,column:0}));
          const readCell = engine.stagedCell.bind(engine);
          let reads = 0;
          engine.stagedCell = (...args) => { reads++; return readCell(...args); };
          for (let round=1; round<=3; round++) {
            const before = reads;
            const start = performance.now();
            const styles = await engine.conditionalStagedViewport(engine.head,engine.metadata,sheet.id,rules,positions);
            const elapsedMs = performance.now()-start;
            if (styles.length!==1000 || styles.some((style,row)=>!!style.bold!==(row>0))) throw new Error('Incorrect appearance');
            if (reads-before!==1000) throw new Error('Unexpected staged cell reads');
            results.push({cacheBytes:cache,round,elapsedMs,verified:styles.length,stagedCellReads:reads-before});
          }
        } finally { engine.close(); }
      }
      self.postMessage({results});
    })().catch(error=>self.postMessage({error:String(error),stack:error.stack}));
  `,
    }),
  );
  await page.goto('http://127.0.0.1:5173/');
  report.result = await page.evaluate(
    () =>
      new Promise((resolve, reject) => {
        const worker = new Worker('/conditional-staged-benchmark.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker timed out'));
        }, 120000);
        worker.onerror = (event) => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error(event.message));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timer);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  if (report.result.error) throw new Error(report.result.error);
  report.samples = report.result.results;
  delete report.result;
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report));
} catch (error) {
  report.error = String(error);
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  throw error;
} finally {
  await browser.close();
}

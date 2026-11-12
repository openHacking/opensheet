import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, browserName = 'chromium', mode = 'tail'] = process.argv.slice(2);
const launchers = { chromium, firefox, webkit };
if (!output || !launchers[browserName] || !['head', 'tail'].includes(mode))
  throw new Error(
    'Usage: node scripts/axis-page-delete-small-paired.mjs NEW_REPORT [chromium|firefox|webkit] [head|tail]',
  );
const enginePath = resolve('packages/core/src/engine.ts');
const source = readFileSync(enginePath, 'utf8');
const needle = `      if (deletion && !remainingBuckets.has(payload.axis === 'row' ? +br : +bc)) {
        this.changed.add(address);
        delete next.blocks[address];
        continue;
      }
`;
if (source.split(needle).length !== 2) throw new Error('Cannot reconstruct control');
const scanStart = source.indexOf('    const remainingBuckets = new Set<number>();');
const scanEnd = source.indexOf('    metadata.snapshot = transformed;', scanStart);
if (scanStart < 0 || scanEnd < 0) throw new Error('Cannot reconstruct survivor-index control');
const control =
  source.replace(needle, '').slice(0, scanStart) + source.replace(needle, '').slice(scanEnd);
const variants = ['control', 'current'];
const paths = variants.map((name) => `apps/playground/public/axis-delete-small-${name}.mjs`);
if (paths.some(existsSync)) throw new Error('Refusing to overwrite a benchmark bundle');
const report = {
  startedAt: new Date().toISOString(),
  browserName,
  mode,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  source: sourceFingerprint(),
  scriptSha256: createHash('sha256')
    .update(readFileSync('scripts/axis-page-delete-small-paired.mjs'))
    .digest('hex'),
  scope:
    'Real Worker and IndexedDB, 20,000 numeric rows, delete one row at the head or tail. Three alternating fresh-context pairs; control removes both the survivor scan and whole-page shortcut. Counts engine block calls, not physical IndexedDB reads. Setup, undo/redo and all-value verification excluded from timed deletion. Bundled source variants, no renderer, input latency, peak memory, independent hardware or power-mode record.',
  bundles: {},
  samples: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const created = [];
let browser;
try {
  for (const [index, name] of variants.entries()) {
    const bundle = await build({
      stdin: {
        contents: `export { Engine } from ${JSON.stringify(enginePath)}; export { createSnapshot } from ${JSON.stringify(resolve('packages/core/src/model.ts'))}; export { Database } from ${JSON.stringify(resolve('packages/core/src/storage.ts'))};`,
        resolveDir: process.cwd(),
        loader: 'ts',
      },
      plugins: [
        {
          name: 'variant',
          setup(builder) {
            builder.onLoad({ filter: /core\/src\/engine\.ts$/ }, () => ({
              contents: name === 'control' ? control : source,
              loader: 'ts',
              resolveDir: resolve('packages/core/src'),
            }));
          },
        },
      ],
      bundle: true,
      platform: 'browser',
      format: 'esm',
      write: false,
      minify: true,
      logLevel: 'silent',
    });
    const bytes = bundle.outputFiles[0].contents;
    writeFileSync(paths[index], bytes, { flag: 'wx' });
    created.push(paths[index]);
    report.bundles[name] = createHash('sha256').update(bytes).digest('hex');
  }
  browser = await launchers[browserName].launch();
  report.browser = browser.version();
  for (let round = 1; round <= 3; round++)
    for (const variant of round % 2 ? variants : [...variants].reverse()) {
      const context = await browser.newContext();
      try {
        const page = await context.newPage();
        await page.route('**/axis-delete-small.worker.js', (route) =>
          route.fulfill({
            contentType: 'text/javascript',
            body: `
import { Engine, createSnapshot, Database } from '/axis-delete-small-${variant}.mjs';
(async () => {
  const snapshot=createSnapshot({sheets:[{name:'Data',rows:20000,columns:1}]});
  const sheet=snapshot.sheets[0];
  for(let row=0;row<20000;row++) sheet.cells[row+':0']={rowId:row,columnId:0,input:{type:'number',value:row}};
  const db=await Database.open('axis-paired-'+crypto.randomUUID());
  const engine=new Engine(db,snapshot.workbookId,0);
  await engine.open(snapshot);
  try {
    const block=engine.block.bind(engine); let blockCalls=0;
    engine.block=(...args)=>{blockCalls++;return block(...args);};
    const started=performance.now();
    const deletedRow=${mode === 'head' ? 0 : 19999};
    await engine.execute([{type:'core.axis.delete',payload:{sheetId:sheet.id,axis:'row',index:deletedRow,count:1}}]);
    const elapsedMs=performance.now()-started, deletionBlockCalls=blockCalls;
    const verify=async(count, removed)=>{
      const order=engine.metadata.snapshot.sheets[0].rowOrder;
      if(order.length!==count) throw new Error('Wrong row count');
      for(let start=0;start<count;start+=10000){
        const end=Math.min(count,start+10000);
        const result=await engine.read(sheet.id,{startRow:start,endRow:end,startColumn:0,endColumn:1},false,false);
        for(let row=start;row<end;row++) {
          const id=order[row], value=result.cells[id+':0']?.input;
          const expected=removed && row>=deletedRow ? row+1 : row;
          if(id!==expected||value?.type!=='number'||value.value!==expected) throw new Error('Wrong value '+row);
        }
      }
    };
    await verify(19999,true);
    await engine.history(false); await verify(20000,false);
    await engine.history(true); await verify(19999,true);
    self.postMessage({elapsedMs,blockCalls:deletionBlockCalls,verified:59998});
  } finally { engine.close(); }
})().catch(error=>self.postMessage({error:String(error),stack:error.stack}));`,
          }),
        );
        await page.goto('http://127.0.0.1:5173/');
        const result = await page.evaluate(
          () =>
            new Promise((resolve, reject) => {
              const worker = new Worker('/axis-delete-small.worker.js', { type: 'module' });
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
        if (result.error) throw new Error(result.error);
        report.samples.push({ round, variant, ...result });
        save();
        console.log(JSON.stringify(report.samples.at(-1)));
      } finally {
        await context.close();
      }
    }
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  save();
  throw error;
} finally {
  await browser?.close();
  for (const path of created) unlinkSync(path);
}

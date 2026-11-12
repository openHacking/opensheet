import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, browserName = 'chromium'] = process.argv.slice(2);
const launchers = { chromium, firefox, webkit };
if (!output || !launchers[browserName])
  throw new Error(
    'Usage: node scripts/recalculation-zero-scratch-paired.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const enginePath = resolve('packages/core/src/engine.ts');
const source = readFileSync(enginePath, 'utf8');
const needle = 'this.calculator.readSession(16 * 1024)';
if (source.split(needle).length !== 2) throw new Error('Cannot reconstruct zero-scratch control');
const paths = ['control', 'current'].map(
  (name) => `apps/playground/public/zero-scratch-${name}.mjs`,
);
if (paths.some(existsSync)) throw new Error('Refusing to overwrite benchmark bundles');
const report = {
  startedAt: new Date().toISOString(),
  browserName,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  source: sourceFingerprint(),
  bundles: {},
  samples: [],
  scope:
    'Real Worker and IndexedDB, zero persistent cache, 512 formulas referencing the same 32 scalar cells. Three alternating fresh-context pairs. Control reconstructs the previous zero-scratch read; setup excluded, every result checked. Source cell calls are not physical storage reads. Development-server source bundles, no production UI, peak memory, other devices or reference performance acceptance.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const created = [];
let browser;
try {
  for (const [index, name] of ['control', 'current'].entries()) {
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
              contents:
                name === 'control'
                  ? source.replace(needle, 'this.calculator.readSession()')
                  : source,
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
    for (const variant of round % 2 ? ['control', 'current'] : ['current', 'control']) {
      const context = await browser.newContext();
      try {
        const page = await context.newPage();
        await page.route('**/zero-scratch.worker.js', (route) =>
          route.fulfill({
            contentType: 'text/javascript',
            body: `
import {Engine,createSnapshot,Database} from '/zero-scratch-${variant}.mjs';
(async () => {
 const snapshot=createSnapshot({sheets:[{name:'Source',rows:512,columns:1},{name:'Output',rows:512,columns:1}]});
 const [source,output]=snapshot.sheets;
 for(let row=0;row<32;row++) source.cells[row+':0']={rowId:row,columnId:0,input:{type:'number',value:1}};
 for(let row=0;row<512;row++) output.cells[row+':0']={rowId:row,columnId:0,input:{type:'formula',expression:'SUM(Source!$A$1:$A$32)+'+row}};
 const db=await Database.open('zero-scratch-'+crypto.randomUUID());
 const engine=new Engine(db,snapshot.workbookId,0);
 await engine.open(snapshot);
 try {
  const cell=engine.cell.bind(engine); let sourceReads=0;
  engine.cell=(sheet,...rest)=>{if(sheet===source.id) sourceReads++;return cell(sheet,...rest);};
  const start=performance.now();
  const result=await engine.read(output.id,{startRow:0,endRow:512,startColumn:0,endColumn:1},true,false);
  const elapsedMs=performance.now()-start;
  for(let row=0;row<512;row++) if(result.calculated[row+':0']!==32+row) throw new Error('Wrong result '+row);
  if(sourceReads!==${variant === 'control' ? 16384 : 32}) throw new Error('Unexpected source calls '+sourceReads);
  self.postMessage({elapsedMs,sourceReads,verified:512});
 } finally { engine.close(); }
})().catch(error=>self.postMessage({error:String(error),stack:error.stack}));`,
          }),
        );
        await page.goto('http://127.0.0.1:5173/');
        const result = await page.evaluate(
          () =>
            new Promise((resolveResult, reject) => {
              const worker = new Worker('/zero-scratch.worker.js', { type: 'module' });
              const timer = setTimeout(() => {
                worker.terminate();
                reject(new Error('Worker timeout'));
              }, 120000);
              worker.onerror = (e) => {
                clearTimeout(timer);
                worker.terminate();
                reject(new Error(e.message));
              };
              worker.onmessage = ({ data }) => {
                clearTimeout(timer);
                worker.terminate();
                resolveResult(data);
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

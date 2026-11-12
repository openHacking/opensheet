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
  throw new Error('Usage: node scripts/read-yield-paired.mjs NEW_REPORT [chromium|firefox|webkit]');
const enginePath = resolve('packages/core/src/engine.ts');
const source = readFileSync(enginePath, 'utf8');
const needle =
  'if (++visited % 1024 === 0) {\n          await new Promise<void>((resolve) => setTimeout(resolve, 0));\n          this.check();\n        }';
if (source.split(needle).length !== 2) throw new Error('Cannot reconstruct old scheduling');
const paths = ['control', 'current'].map((name) => `apps/playground/public/read-yield-${name}.mjs`);
if (paths.some(existsSync)) throw new Error('Refusing to overwrite bundles');
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
    'Three alternating fresh-context pairs, real Worker and IndexedDB; 50 repeated blank 16x16 viewport reads and one 2048x1 narrow read. Control reconstructs the prior coordinate-based yield condition. Counts Worker setTimeout calls, verifies all returned cells, excludes setup, production UI, concurrent user input, peak memory and other devices.',
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
                  ? source.replace(
                      needle,
                      'if ((row * sh.columnOrder.length + column) % 1024 === 0)\n          await new Promise<void>((resolve) => setTimeout(resolve, 0));',
                    )
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
        await page.route('**/read-yield.worker.js', (route) =>
          route.fulfill({
            contentType: 'text/javascript',
            body: `
import {Engine,createSnapshot,Database} from '/read-yield-${variant}.mjs';
(async()=>{
 const snapshot=createSnapshot({sheets:[{name:'Data',rows:2048,columns:16}]});
 const sheet=snapshot.sheets[0],db=await Database.open('read-yield-'+crypto.randomUUID());
 const engine=new Engine(db,snapshot.workbookId);await engine.open(snapshot);
 try {
  let timers=0;const oldSetTimeout=self.setTimeout;self.setTimeout=(...args)=>{timers++;return oldSetTimeout(...args);};
  const small={startRow:0,endRow:16,startColumn:0,endColumn:16};
  const begin=performance.now();let checked=0;
  for(let i=0;i<50;i++){const r=await engine.read(sheet.id,small,true,false);const values=Object.values(r.calculated);if(values.length!==256||values.some(x=>x!==null))throw Error('small mismatch');checked+=values.length;}
  const smallMs=performance.now()-begin,smallTimers=timers;timers=0;
  const start=performance.now();const r=await engine.read(sheet.id,{startRow:0,endRow:2048,startColumn:1,endColumn:2},true,false);
  const largeMs=performance.now()-start,largeTimers=timers;const values=Object.values(r.calculated);
  if(values.length!==2048||values.some(x=>x!==null))throw Error('large mismatch');
  self.postMessage({smallMs,smallTimers,largeMs,largeTimers,verified:checked+values.length});
 }finally{engine.close();}
})().catch(error=>self.postMessage({error:String(error),stack:error.stack}));`,
          }),
        );
        await page.goto('http://127.0.0.1:5173/');
        const result = await page.evaluate(
          () =>
            new Promise((resolveResult, reject) => {
              const worker = new Worker('/read-yield.worker.js', { type: 'module' });
              const timer = setTimeout(() => {
                worker.terminate();
                reject(Error('Worker timeout'));
              }, 120000);
              worker.onerror = (e) => {
                clearTimeout(timer);
                worker.terminate();
                reject(Error(e.message));
              };
              worker.onmessage = ({ data }) => {
                clearTimeout(timer);
                worker.terminate();
                resolveResult(data);
              };
            }),
        );
        if (result.error) throw Error(result.error);
        const expectedSmall = variant === 'control' ? 50 : 0,
          expectedLarge = variant === 'control' ? 0 : 2;
        if (result.smallTimers !== expectedSmall || result.largeTimers !== expectedLarge)
          throw Error('Unexpected yield counts ' + JSON.stringify(result));
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

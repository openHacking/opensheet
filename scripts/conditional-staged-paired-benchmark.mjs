import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs';
import { arch, cpus, platform, totalmem } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2],
  browserName = process.argv[3] ?? 'chromium',
  mode = process.argv[4] ?? 'scalar';
const launchers = { chromium, firefox, webkit };
if (!output || !launchers[browserName] || !['scalar', 'page'].includes(mode))
  throw new Error(
    'Usage: node scripts/conditional-staged-paired-benchmark.mjs NEW_REPORT [chromium|firefox|webkit] [scalar|page]',
  );
const enginePath = resolve('packages/core/src/engine.ts');
const source = readFileSync(enginePath, 'utf8');
const needle =
  mode === 'scalar'
    ? `const scalarMemo = checked.some((rule) => rule.condition.kind === 'formula')\n      ? new Map<number, CellRecord | undefined>()\n      : undefined;`
    : 'const viewportPage: StagedReadPage = {};';
if (source.split(needle).length !== 2) throw new Error('Cannot reconstruct the no-reuse control');
const control = source.replace(
  needle,
  mode === 'scalar'
    ? 'const scalarMemo = undefined as Map<number, CellRecord | undefined> | undefined;'
    : 'const viewportPage: StagedReadPage | undefined = undefined;',
);
const names = ['control', 'current'];
const paths = names.map((name) => `apps/playground/public/conditional-paired-${name}.mjs`);
if (paths.some(existsSync)) throw new Error('Refusing to overwrite a benchmark bundle');
const report = {
  startedAt: new Date().toISOString(),
  engine: browserName,
  mode,
  source: sourceFingerprint(),
  host: { platform: platform(), arch: arch(), cpu: cpus()[0].model, memoryBytes: totalmem() },
  scope: `Three alternating fresh-context pairs, real Worker and IndexedDB, 1,000 numeric cells and one relative formula predicate, zero cache reservation. Control is the same source with only the per-viewport ${mode === 'scalar' ? 'scalar map' : 'page reuse'} disabled. Setup and all-value checking are outside timing. No production bundle, renderer, frame timing, peak process memory, other devices or formula-heavy values.`,
  bundles: {},
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const created = [];
let browser;
try {
  for (const [index, name] of names.entries()) {
    const bundle = await build({
      stdin: {
        contents: `export { Engine } from ${JSON.stringify(enginePath)}; export { createSnapshot } from ${JSON.stringify(resolve('packages/core/src/model.ts'))}; export { Database } from ${JSON.stringify(resolve('packages/core/src/storage.ts'))};`,
        resolveDir: process.cwd(),
        loader: 'ts',
      },
      plugins: [
        {
          name: 'engine-variant',
          setup: (builder) =>
            builder.onLoad({ filter: /core\/src\/engine\.ts$/ }, () => ({
              contents: name === 'control' ? control : source,
              loader: 'ts',
              resolveDir: resolve('packages/core/src'),
            })),
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
  for (let round = 1; round <= 3; round++) {
    for (const variant of round % 2 ? names : [...names].reverse()) {
      const context = await browser.newContext();
      try {
        const page = await context.newPage();
        await page.route('**/conditional-paired.worker.js', (route) =>
          route.fulfill({
            contentType: 'text/javascript',
            body: `
          import { Engine, createSnapshot, Database } from '/conditional-paired-${variant}.mjs';
          (async () => {
            const snapshot = createSnapshot({sheets:[{name:'Data',rows:1000,columns:1}]});
            const sheet = snapshot.sheets[0];
            for(let row=0;row<1000;row++) sheet.cells[row+':0']={rowId:row,columnId:0,input:{type:'number',value:row}};
            const db=await Database.open('conditional-paired-'+crypto.randomUUID());
            const engine=new Engine(db,snapshot.workbookId,0);
            await engine.open(snapshot);
            try {
              const rules=[{id:'positive',range:{startRow:0,endRow:1000,startColumn:0,endColumn:1},condition:{kind:'formula',expression:'A1>0'},style:{bold:true},stopIfTrue:false}];
              const positions=Array.from({length:1000},(_,row)=>({row,column:0}));
              const read=engine.stagedCell.bind(engine); let reads=0;
              engine.stagedCell=(...args)=>{reads++;return read(...args);};
              const load=engine.block.bind(engine); let pageReads=0;
              engine.block=(...args)=>{pageReads++;return load(...args);};
              const start=performance.now();
              const styles=await engine.conditionalStagedViewport(engine.head,engine.metadata,sheet.id,rules,positions);
              const elapsedMs=performance.now()-start;
              if(styles.length!==1000||styles.some((style,row)=>!!style.bold!==(row>0))) throw new Error('Incorrect styles');
              self.postMessage({elapsedMs,reads,pageReads,verified:styles.length});
            } finally {engine.close();}
          })().catch(error=>self.postMessage({error:String(error),stack:error.stack}));
        `,
          }),
        );
        await page.goto('http://127.0.0.1:5173/');
        const result = await page.evaluate(
          () =>
            new Promise((resolve, reject) => {
              const worker = new Worker('/conditional-paired.worker.js', { type: 'module' });
              const timer = setTimeout(() => {
                worker.terminate();
                reject(new Error('Worker timed out'));
              }, 60000);
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
        if (result.reads !== (mode === 'scalar' && variant === 'control' ? 2000 : 1000))
          throw new Error('Unexpected read count');
        if (mode === 'page' && result.pageReads !== (variant === 'control' ? 1000 : 16))
          throw new Error('Unexpected page count');
        report.samples.push({ round, variant, ...result });
        writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
        console.log(JSON.stringify(report.samples.at(-1)));
      } finally {
        await context.close();
      }
    }
  }
} catch (error) {
  report.error = String(error);
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  throw error;
} finally {
  await browser?.close();
  for (const path of created) unlinkSync(path);
}

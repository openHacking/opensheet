import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new JSON report path');
const modulePath = resolve('packages/core/src/validation-batch.ts');
const source = readFileSync(modulePath, 'utf8');
const needle = `    if (
      cell.row < bounds.startRow ||
      cell.row >= bounds.endRow ||
      cell.column < bounds.startColumn ||
      cell.column >= bounds.endColumn
    ) {
      const previous = Math.floor(scanned / CHECKPOINT_INTERVAL);
      scanned += checkedRules.length;
      if (Math.floor(scanned / CHECKPOINT_INTERVAL) !== previous) yield CHECKPOINT;
      continue;
    }
`;
if (source.split(needle).length !== 2) throw new Error('Cannot reconstruct full-scan control');
const control = source.replace(needle, '');
const variants = {};
for (const [name, contents] of Object.entries({ control, current: source })) {
  const bundle = await build({
    stdin: {
      contents: `import { checkStagedValidation } from ${JSON.stringify(modulePath)}; globalThis.__validationBench = checkStagedValidation;`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /core\/src\/validation-batch\.ts$/ }, () => ({
            contents,
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
  });
  variants[name] = bundle.outputFiles[0].text;
}
const hash = (value) => createHash('sha256').update(value).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  scriptSha256: hash(readFileSync('scripts/validation-bounds-paired.mjs')),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  bundles: Object.fromEntries(Object.entries(variants).map(([name, text]) => [name, hash(text)])),
  scope:
    'Headless real Worker, local development server, three alternating fresh-context pairs per browser and workload. Control disables only the bounding-box skip, retaining bounds calculation. Workloads: 65,536 positions and no rules; or 1,000 positions and 900 disjoint or covering numeric rules. Includes position copying, rule preparation, in-memory readers and asynchronous checkpoints; no persistence, UI, peak memory, independent devices or power-mode record. Uses actual worker clocks, no forced cancellation. Verifies full report and read count.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const entry = { engine, version: browser.version(), samples: [] };
  report.results.push(entry);
  try {
    for (const fixture of ['positions', 'disjoint', 'covered'])
      for (let round = 1; round <= 3; round++)
        for (const variant of round % 2 ? ['control', 'current'] : ['current', 'control']) {
          const context = await browser.newContext();
          try {
            const page = await context.newPage();
            await page.route('**/validation-preflight.worker.js', (route) =>
              route.fulfill({
                contentType: 'text/javascript',
                body: `${variants[variant]}
              const positions = Array.from({length:${fixture === 'positions' ? 65536 : 1000}}, (_, row) => ({row,column:0}));
              const rule = {kind:'number', operator:'gte', first:1, allowBlank:false, policy:'reject'};
              const rules = ${fixture === 'positions' ? '[]' : `Array.from({length:900},(_,i)=>({id:String(i),range:{startRow:${fixture === 'covered' ? 0 : 1000},endRow:${fixture === 'covered' ? 1000 : 1001},startColumn:0,endColumn:1},rule}))`};
              let reads=0;
              const start=performance.now();
              globalThis.__validationBench(rules,positions,async()=>{reads++;return 1;}).then(result=>self.postMessage({elapsedMs:performance.now()-start,reads,result}),error=>self.postMessage({error:String(error)}));`,
              }),
            );
            await page.goto('http://127.0.0.1:5173/');
            const result = await page.evaluate(
              () =>
                new Promise((resolve, reject) => {
                  const worker = new Worker('/validation-preflight.worker.js', { type: 'module' });
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
            if (
              result.error ||
              result.reads !== (fixture === 'covered' ? 1000 : 0) ||
              JSON.stringify(result.result) !==
                JSON.stringify({ rejected: 0, warnings: 0, details: [], truncated: false })
            )
              throw new Error(`Unexpected result: ${JSON.stringify(result)}`);
            entry.samples.push({ fixture, round, variant, elapsedMs: result.elapsedMs });
            save();
            console.log(JSON.stringify({ engine, ...entry.samples.at(-1) }));
          } finally {
            await context.close();
          }
        }
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();

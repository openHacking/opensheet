import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new JSON report path');
const path = resolve('packages/core/src/range.ts');
const source = readFileSync(path, 'utf8');
const start = source.indexOf('  private write<T>(');
const end = source.indexOf('  setValues(values:', start);
if (start < 0 || end < 0) throw new Error('Cannot reconstruct control');
const control =
  source.slice(0, start) +
  `  private write<T>(matrix: T[][], map: (value: T) => CellInput) {
    const b = this.bounds;
    assert(matrix.length === b.endRow - b.startRow && matrix.every((row) => row.length === b.endColumn - b.startColumn), 'INVALID_ARGUMENT', 'Matrix dimensions must match the range');
    return this.sheet.workbook.execute({type:'core.cells.set',payload:{sheetId:this.sheet.id,cells:matrix.flatMap((row,i)=>row.map((value,j)=>({row:b.startRow+i,column:b.startColumn+j,input:map(value)})))}});
  }
` +
  source.slice(end);
const bundles = {};
for (const [name, contents] of Object.entries({ control, current: source })) {
  const bundle = await build({
    stdin: {
      contents: `export { Range } from ${JSON.stringify(path)};`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'range-variant',
        setup(builder) {
          builder.onLoad({ filter: /core\/src\/range\.ts$/ }, () => ({
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
  bundles[name] = bundle.outputFiles[0].text;
}
const hash = (data) => createHash('sha256').update(data).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  scriptSha256: hash(readFileSync('scripts/matrix-write-browser-paired.mjs')),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  bundles: Object.fromEntries(Object.entries(bundles).map(([name, body]) => [name, hash(body)])),
  scope:
    'Browser main-thread Range matrix-to-command construction only; 100,000 dense numeric or formula inputs. Three fresh contexts per browser, four alternating pairs per fixture and context; every coordinate and input checked outside timing. Control reconstructs old flatMap path, which silently skips sparse entries. Excludes real Worker, transfer, schema validation, persistence, actual paint/input latency, peak memory, other hardware and power-mode record; sequential samples are sensitive to JIT/GC.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const entry = { engine, version: browser.version(), rounds: [] };
  report.results.push(entry);
  try {
    for (let round = 0; round < 3; round++) {
      const context = await browser.newContext();
      try {
        const page = await context.newPage();
        for (const [name, body] of Object.entries(bundles))
          await page.route(`**/matrix-range-${name}.mjs`, (route) =>
            route.fulfill({ contentType: 'text/javascript', body }),
          );
        await page.goto('http://127.0.0.1:5173/');
        const samples = await page.evaluate(async (round) => {
          const variants = {
            control: (await import('/matrix-range-control.mjs')).Range,
            current: (await import('/matrix-range-current.mjs')).Range,
          };
          const rows = 1000,
            columns = 100,
            out = [];
          for (const fixture of ['numeric', 'formula']) {
            const matrix = Array.from({ length: rows }, (_, r) =>
              Array.from({ length: columns }, (_, c) =>
                fixture === 'numeric' ? r * columns + c : '=A1+1',
              ),
            );
            for (let pair = 0; pair < 4; pair++)
              for (const variant of (round + pair) % 2
                ? ['current', 'control']
                : ['control', 'current']) {
                let command;
                const sheet = {
                  id: 'sheet',
                  workbook: {
                    execute(value) {
                      command = value;
                    },
                  },
                };
                const range = new variants[variant](sheet, {
                  startRow: 0,
                  endRow: rows,
                  startColumn: 0,
                  endColumn: columns,
                });
                const start = performance.now();
                if (fixture === 'numeric') range.setValues(matrix);
                else range.setFormulas(matrix);
                const elapsedMs = performance.now() - start;
                if (command.payload.cells.length !== rows * columns) throw new Error('Wrong count');
                for (let i = 0; i < rows * columns; i++) {
                  const cell = command.payload.cells[i];
                  if (
                    cell.row !== Math.floor(i / columns) ||
                    cell.column !== i % columns ||
                    (fixture === 'numeric'
                      ? cell.input.type !== 'number' || cell.input.value !== i
                      : cell.input.type !== 'formula' || cell.input.expression !== 'A1+1')
                  )
                    throw new Error(`Wrong cell ${fixture}/${i}`);
                }
                out.push({ fixture, pair: pair + 1, variant, elapsedMs });
              }
          }
          return out;
        }, round);
        entry.rounds.push({ round: round + 1, samples });
        save();
        console.log(JSON.stringify({ engine, round: round + 1, samples: samples.length }));
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

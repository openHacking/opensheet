import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium', workload = 'uniform'] = process.argv.slice(2);
if (
  !output ||
  !['chromium', 'firefox', 'webkit'].includes(engine) ||
  !['uniform', 'mixed'].includes(workload)
)
  throw new Error(
    'Usage: node scripts/structural-edit-benchmark.mjs NEW_REPORT [chromium|firefox|webkit] [uniform|mixed]',
  );
const rows = 1024;
const columns = 32;
const file = createWorkbookFile({
  sheets: [
    { name: 'Inputs', rows: 99_999, columns },
    { name: 'Formulas', rows, columns },
  ],
});
file.sheets[0].blocks.push({
  row: 0,
  column: 0,
  data: {
    numbers:
      workload === 'mixed'
        ? { positions: [0, 1], values: [41, 7] }
        : { positions: [0], values: [41] },
  },
});
const mixed = (index) => {
  const constant = index + 1;
  switch (index % 4) {
    case 0:
      return [`Inputs!$A$1+${constant}`, 41 + constant];
    case 1:
      return [`Inputs!B1*2+${constant}`, 14 + constant];
    case 2:
      return [`SUM(Inputs!$A1:$B1)+${constant}`, 48 + constant];
    default:
      return [`IF(Inputs!A$1>Inputs!$B$1,${constant},0)`, constant];
  }
};
const samples = [];
for (let row = 0; row < rows / 64; row++) {
  const start = row * 2048;
  if (workload === 'mixed') {
    for (let i = 0; i < 4; i++)
      samples.push([`${String.fromCharCode(65 + i)}${row * 64 + 1}`, mixed(start + i)[1]]);
    samples.push([`AF${(row + 1) * 64}`, mixed(start + 2047)[1]]);
  }
  file.sheets[1].blocks.push({
    row,
    column: 0,
    data: {
      formulas: {
        positions: Array.from({ length: 2048 }, (_, i) => i),
        values:
          workload === 'mixed'
            ? Array.from({ length: 2048 }, (_, i) => mixed(start + i)[0])
            : Array(2048).fill('Inputs!$A$1+1'),
      },
    },
  });
}
if (!samples.length) samples.push(['A1', 42], ['AF1024', 42]);
const report = {
  engine,
  workload,
  cpu: cpus()[0]?.model,
  os: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  startedAt: new Date().toISOString(),
  formulaCells: rows * columns,
  formulaPages: rows / 64,
  sourceRows: 99_999,
  verifiedFormulaCellsPerCheck: samples.length,
  rounds: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  report.browser = browser.version();
  for (let round = 1; round <= 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5186/#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(
        async ({ fixture, samples }) => {
          const app = window.opensheet;
          await app.load(fixture);
          let book = app.getWorkbook();
          const check = async (row) => {
            const [inputs, formulas] = book.getSheets();
            if ((await inputs.range(`A${row}`).getValues())[0][0] !== 41)
              throw new Error('Source moved to the wrong row');
            for (const [address, expected] of samples) {
              const actual = (await formulas.range(address).getValues())[0][0];
              if (actual !== expected)
                throw new Error(
                  `Structural edit changed ${address}: ${JSON.stringify(actual)} != ${expected}`,
                );
            }
          };
          await check(1);
          const revision = book.revision;
          const start = performance.now();
          await book.getSheets()[0].insertRows(0, 1);
          const insertMs = performance.now() - start;
          if (book.revision !== revision + 1) throw new Error('Missing committed revision');
          await check(2);
          await app.open(book.id);
          book = app.getWorkbook();
          await check(2);
          if (!(await book.undo())) throw new Error('Missing undo');
          await check(1);
          if (!(await book.redo())) throw new Error('Missing redo');
          await check(2);
          return { insertMs, encodedBytes: (await book.storageStats()).bytes };
        },
        { fixture: file, samples },
      );
      report.rounds.push({ round, ...result });
      console.log(JSON.stringify({ round, ...result }));
    } catch (error) {
      report.rounds.push({ round, error: String(error) });
      throw error;
    } finally {
      save();
      await context.close();
    }
  }
} finally {
  await browser.close();
}

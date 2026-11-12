import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2],
  url = process.argv[3];
const appMode = process.argv.includes('--app');
if (!output || !url)
  throw new Error('Usage: node scripts/navigation-benchmark.mjs NEW_REPORT DEV_SERVER_URL');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  localBuild: appMode ? localBuildFingerprint() : null,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  url,
  scope: appMode
    ? 'Headless app real Worker, 100k rows in the existing scene workbook; five alternating timed reads after warmup for a 10k contiguous run and an 85k-row sparse gap. Three additional pairs time synthetic keydown through selection and grid.ready; three bursts time 32 queued keydowns and check every selection. Every fixture input and destination checked. Source and local build fingerprints do not independently verify served assets. Excludes native keyboard latency, pixel inspection, peak memory and other machines.'
    : 'Headless development-server real Worker, one 100k-row one-column workbook; five alternating timed reads after warmup for a 10k contiguous run and an 85k-row sparse gap. Every destination checked. Excludes user input, selection paint, production build, peak memory and other machines.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), samples: [] };
  report.results.push(result);
  try {
    const page = await browser.newPage();
    await page.goto(url);
    if (appMode)
      await page.waitForFunction(() => {
        const select = document.querySelector('#scene-select');
        return select && !select.disabled;
      });
    const measurement = await page.evaluate(
      async ({ moduleUrl, appMode }) => {
        const book = appMode
          ? window.opensheet.getWorkbook()
          : await (
              await import(/* @vite-ignore */ moduleUrl)
            ).createWorkbook({ sheets: [{ name: 'Probe', rows: 100000, columns: 1 }] });
        try {
          const sheet = book.getSheets()[0];
          if (appMode && sheet.rowCount < 100000)
            await sheet.insertRows(sheet.rowCount, 100000 - sheet.rowCount);
          await sheet.range('A1:A10000').setValues(Array.from({ length: 10000 }, (_, i) => [i]));
          await sheet.range('A100000').setValues([[1]]);
          const denseValues = await sheet.range('A1:A10000').getValues();
          if (denseValues.length !== 10000 || denseValues.some((row, index) => row[0] !== index))
            throw new Error('Wrong dense fixture values');
          if ((await sheet.range('A100000').getValues())[0][0] !== 1)
            throw new Error('Wrong sparse fixture endpoint');
          for (let startRow = 15000; startRow < 99999; startRow += 10000) {
            const cells = await sheet
              .range({
                startRow,
                endRow: Math.min(99999, startRow + 10000),
                startColumn: 0,
                endColumn: 1,
              })
              .getInputs();
            if (cells.some((row) => row[0].type !== 'blank'))
              throw new Error('Wrong sparse fixture gap');
          }
          const run = async (name) => {
            const start = performance.now();
            const result = await book.navigateData(
              sheet.id,
              name === 'dense' ? 0 : 15000,
              0,
              'row',
              1,
            );
            const ms = performance.now() - start;
            if (
              result.row !== (name === 'dense' ? 9999 : 99999) ||
              result.column !== 0 ||
              result.revision !== book.revision
            )
              throw new Error('Wrong navigation result');
            return ms;
          };
          await run('dense');
          await run('sparse');
          const samples = [];
          for (let i = 0; i < 5; i++) {
            const pair = {};
            for (const name of i % 2 ? ['sparse', 'dense'] : ['dense', 'sparse'])
              pair[name] = await run(name);
            samples.push(pair);
          }
          const uiSamples = [];
          const burstSamples = [];
          if (appMode) {
            const app = window.opensheet,
              grid = app.getGrid();
            if (!grid) throw new Error('Missing grid');
            const uiRun = async (name) => {
              const row = name === 'dense' ? 0 : 15000;
              const target = name === 'dense' ? 9999 : 99999;
              app.selection.set({
                sheetId: sheet.id,
                startRow: row,
                endRow: row + 1,
                startColumn: 0,
                endColumn: 1,
              });
              grid.focus();
              await grid.ready();
              const start = performance.now();
              const event = new KeyboardEvent('keydown', {
                key: 'ArrowDown',
                ctrlKey: true,
                bubbles: true,
                cancelable: true,
              });
              grid.element.dispatchEvent(event);
              if (!event.defaultPrevented) throw new Error('Navigation key was not handled');
              while (app.selection.get()?.startRow !== target) {
                if (performance.now() - start > 10000) throw new Error('Navigation UI timed out');
                await new Promise((resolve) => requestAnimationFrame(resolve));
              }
              await grid.ready();
              if (app.selection.get()?.startColumn !== 0) throw new Error('Wrong UI column');
              return performance.now() - start;
            };
            await uiRun('dense');
            await uiRun('sparse');
            for (let i = 0; i < 3; i++) {
              const pair = {};
              for (const name of i % 2 ? ['sparse', 'dense'] : ['dense', 'sparse'])
                pair[name] = await uiRun(name);
              uiSamples.push(pair);
            }
            await sheet
              .range('B1:B65')
              .setValues(Array.from({ length: 65 }, (_, i) => [i % 2 === 0 ? i : null]));
            const burstInputs = await sheet.range('B1:B65').getInputs();
            if (
              burstInputs.some((row, i) =>
                i % 2 === 0
                  ? row[0].type !== 'number' || row[0].value !== i
                  : row[0].type !== 'blank',
              )
            )
              throw new Error('Wrong burst fixture');
            const burst = async () => {
              app.selection.set({
                sheetId: sheet.id,
                startRow: 0,
                endRow: 1,
                startColumn: 1,
                endColumn: 2,
              });
              grid.focus();
              await grid.ready();
              const positions = [];
              const stop = app.selection.onChange((selection) =>
                positions.push([selection.startRow, selection.startColumn]),
              );
              const start = performance.now();
              try {
                for (let i = 0; i < 32; i++) {
                  const event = new KeyboardEvent('keydown', {
                    key: 'ArrowDown',
                    ctrlKey: true,
                    bubbles: true,
                    cancelable: true,
                  });
                  grid.element.dispatchEvent(event);
                  if (!event.defaultPrevented) throw new Error('Burst key was not handled');
                }
                while (app.selection.get()?.startRow !== 64) {
                  if (performance.now() - start > 10000)
                    throw new Error('Navigation burst timed out');
                  await new Promise((resolve) => requestAnimationFrame(resolve));
                }
                await grid.ready();
                const ms = performance.now() - start;
                if (
                  positions.length !== 32 ||
                  positions.some(([row, column], i) => row !== (i + 1) * 2 || column !== 1)
                )
                  throw new Error('Navigation burst sequence differs');
                return ms;
              } finally {
                stop();
              }
            };
            await burst();
            for (let i = 0; i < 3; i++) burstSamples.push(await burst());
          }
          return { samples, uiSamples, burstSamples };
        } finally {
          await book.deleteStorage();
        }
      },
      { moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`, appMode },
    );
    Object.assign(result, measurement);
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();

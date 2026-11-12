import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const [
  output,
  engine = 'chromium',
  profile = 'timing',
  comparison = 'current',
  workload = 'blank',
  cache = 'warm',
] = process.argv.slice(2);
if (
  !output ||
  !['chromium', 'firefox', 'webkit'].includes(engine) ||
  !['timing', 'profile'].includes(profile) ||
  !['current', 'paired'].includes(comparison) ||
  (comparison === 'paired' && profile !== 'profile') ||
  !['blank', 'dirty'].includes(workload) ||
  !['warm', 'cold'].includes(cache)
)
  throw new Error(
    'Usage: node scripts/merge-worker-timing.mjs NEW_REPORT [chromium|firefox|webkit] [timing|profile] [current|paired] [blank|dirty] [warm|cold]',
  );
const rows = 20000;
const file = createWorkbookFile({ sheets: [{ name: 'Data', rows, columns: 1 }] });
for (let page = 0; page < Math.ceil(rows / 64); page++) {
  const count = Math.min(64, rows - page * 64),
    positions = Array.from({ length: count }, (_, i) => i * 32);
  const formulaPosition = page === 0 ? 32 : 0;
  file.sheets[0].blocks.push({
    row: page,
    column: 0,
    data: {
      blanks:
        workload === 'dirty'
          ? positions.filter((position) => position !== formulaPosition)
          : positions,
      notes: { positions, values: positions.map((_, i) => `note-${page * 64 + i}`) },
      ...(workload === 'dirty'
        ? {
            formulas: { positions: [formulaPosition], values: ['1+1'] },
            cached: { positions: [formulaPosition], values: [{ type: 'number', value: 2 }] },
          }
        : {}),
    },
  });
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  engine,
  profile,
  comparison,
  workload,
  cache,
  cells: rows,
  scope:
    'Development-server direct Workbook with opt-in Worker operation timing; three fresh contexts per variant, one merge each. Paired mode reconstructs the prior unoptimized page writes in the served engine module. Every note checked before/after, undo, redo and reopen. Timing includes instrumentation and scheduling, not CPU attribution; no production assets, native input or process peak memory.',
  rounds: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  report.browser = browser.version();
  for (let run = 0; run < (comparison === 'paired' ? 6 : 3); run++) {
    const round = Math.floor(run / (comparison === 'paired' ? 2 : 1)) + 1;
    const variant =
      comparison === 'paired'
        ? (round % 2 ? ['control', 'current'] : ['current', 'control'])[run % 2]
        : 'current';
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      if (variant === 'control')
        await page.route('**/packages/core/src/engine.ts*', async (route) => {
          const response = await route.fetch();
          let source = await response.text();
          const begin = source.indexOf(
            '    const mergeDirty = command.type === "core.cells.merge"',
          );
          const end = source.indexOf('    for (const address of addresses) {', begin);
          if (begin < 0 || end < 0) throw new Error('Cannot reconstruct merge page control');
          source = source.slice(0, begin) + source.slice(end);
          const check = source.indexOf('      if (mergeDirty) {');
          const after = source.indexOf(
            '      Object.assign(target.cells, structuredClone(block.cells));',
            check,
          );
          if (check < 0 || after < 0)
            throw new Error('Cannot reconstruct merge dirty scan control');
          source = source.slice(0, check) + source.slice(after);
          const grouping = '    if (target && (!mergeDirty || mergeDirty.size)) {';
          if (source.split(grouping).length !== 2)
            throw new Error('Cannot reconstruct merge page grouping control');
          source = source.replace(grouping, '    if (target) {');
          const groupSkip = '        if (mergeDirty && !mergeDirty.has(address)) continue;\n';
          if (source.split(groupSkip).length !== 2)
            throw new Error('Cannot reconstruct merge page grouping filter');
          source = source.replace(groupSkip, '');
          const skip = '      if (mergeDirty && !mergeDirty.has(address)) continue;\n';
          if (source.split(skip).length !== 2)
            throw new Error('Cannot reconstruct merge page write control');
          source = source.replace(skip, '');
          await route.fulfill({ response, body: source });
        });
      if (profile === 'profile')
        await page.route('**/worker-runtime.ts*', async (route) => {
          const response = await route.fetch();
          await route.fulfill({
            response,
            body:
              (await response.text()) +
              `
          let profileEnabled = false, profileData = {};
          for (const name of ['block','storeBlock','publish']) {
            const original = Engine.prototype[name];
            Engine.prototype[name] = async function(...args) {
              if (!profileEnabled) return original.apply(this,args);
              const start = performance.now();
              try { return await original.apply(this,args); }
              finally { const data = profileData[name] ??= {count:0,ms:0}; data.count++; data.ms += performance.now()-start; }
            };
          }
          addEventListener('message',({data}) => {
            if (data.profileStart) { profileData = {}; profileEnabled = true; }
            if (data.profileStop) { profileEnabled = false; postMessage({mergeProfile:profileData}); }
          });
        `,
          });
        });
      await page.goto('http://127.0.0.1:5173/');
      const result = await page.evaluate(
        async ({ fixture, rows, moduleUrl, profile, variant, workload, cache }) => {
          const { Workbook, openWorkbook } = await import(/* @vite-ignore */ moduleUrl);
          let timing;
          let book = await new Workbook(fixture, {
            ...(cache === 'cold' ? { cacheBytes: 0 } : {}),
            onOperationTiming: (sample) => {
              if (sample.method === 'execute') timing = sample;
            },
          }).ready();
          let sheet = book.getSheets()[0];
          const range = { startRow: 0, endRow: rows, startColumn: 0, endColumn: 1 };
          const check = async (expected) => {
            for (let start = 0; start < rows; start += 8192) {
              const end = Math.min(rows, start + 8192),
                notes = await sheet.range({ ...range, startRow: start, endRow: end }).getNotes();
              if (notes.length !== end - start) throw new Error('Incorrect note count');
              for (let row = start; row < end; row++)
                if (notes[row - start][0] !== `note-${row}`) throw new Error(`Changed note ${row}`);
              if (workload === 'dirty') {
                const formulas = await sheet
                  .range({ ...range, startRow: start, endRow: end })
                  .getFormulas();
                for (let row = start; row < end; row++) {
                  const isFormula = row === 1 || (row > 0 && row % 64 === 0);
                  const wanted = expected === 0 && isFormula ? '1+1' : null;
                  if (formulas[row - start][0] !== wanted)
                    throw new Error(`Changed formula ${row}`);
                }
              }
            }
            if (book.sheetData(sheet.id).merges.length !== expected)
              throw new Error('Incorrect merge state');
          };
          try {
            await check(0);
            timing = undefined;
            const worker = book.worker;
            if (profile === 'profile')
              worker.postMessage({ method: 'cancelRequest', args: [-1], profileStart: true });
            const started = performance.now();
            await sheet.range(range).merge({ discardCoveredValues: workload === 'dirty' });
            const mergeMs = performance.now() - started;
            if (!timing) throw new Error('Missing execution timing');
            const mergeTiming = timing;
            let phases;
            if (profile === 'profile')
              phases = await new Promise((resolve, reject) => {
                const timer = setTimeout(() => reject(new Error('Profile timeout')), 10000);
                worker.addEventListener('message', function onMessage({ data }) {
                  if (!data.mergeProfile) return;
                  worker.removeEventListener('message', onMessage);
                  clearTimeout(timer);
                  resolve(data.mergeProfile);
                });
                worker.postMessage({ method: 'cancelRequest', args: [-1], profileStop: true });
              });
            if (
              profile === 'profile' &&
              (phases?.block?.count !== 313 || phases?.publish?.count !== 1)
            )
              throw new Error('Profile hooks did not observe expected engine methods');
            if (
              profile === 'profile' &&
              (phases?.storeBlock?.count ?? 0) !==
                (variant === 'control' || workload === 'dirty' ? 313 : 0)
            )
              throw new Error('Unexpected page rewrite count');
            await check(1);
            if (!(await book.undo())) throw new Error('Missing undo');
            await check(0);
            if (!(await book.redo())) throw new Error('Missing redo');
            await check(1);
            const id = book.id;
            await book.close();
            book = await openWorkbook(id, cache === 'cold' ? { cacheBytes: 0 } : {});
            sheet = book.getSheets()[0];
            await check(1);
            return {
              mergeMs,
              timing: mergeTiming,
              ...(phases ? { phases } : {}),
              verifiedNotesPerCheck: rows,
              bytes: (await book.storageStats()).bytes,
            };
          } finally {
            await book.deleteStorage();
            book.dispose();
          }
        },
        {
          fixture: file,
          rows,
          moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
          profile,
          variant,
          workload,
          cache,
        },
      );
      report.rounds.push({ round, variant, ...result });
      console.log(JSON.stringify({ round, variant, ...result }));
    } catch (error) {
      report.rounds.push({ round, variant, error: String(error) });
      throw error;
    } finally {
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
      await context.close();
    }
  }
} finally {
  await browser.close();
}

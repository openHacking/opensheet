import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { cpus, totalmem, platform, arch, release } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';
const output = process.argv[2];
if (!output) throw new Error('Provide a new output path');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  host: {
    node: process.version,
    cpu: cpus()[0]?.model,
    memoryBytes: totalmem(),
    platform: platform(),
    arch: arch(),
    release: release(),
    powerMode: 'unrecorded',
  },
  scope:
    'Production preview; 10,000 values in a real Worker, five repeated range display reads per scenario, three fresh contexts per browser. All outputs checked, then reopened and checked. Read timings include Worker, transport and matrix creation, exclude paint; frame intervals sampled separately on the main page, not INP. Worker/process heap unavailable.',
  results: [],
};
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
try {
  const base = 'http://127.0.0.1:5186/',
    hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
  report.artifacts = [];
  for (const file of [
    'index.html',
    ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
  ]) {
    const response = await fetch(new URL(file, base)),
      sha256 = hash(new Uint8Array(await response.arrayBuffer()));
    if (!response.ok || sha256 !== hash(readFileSync(`apps/playground/dist/${file}`)))
      throw new Error(`Build mismatch: ${file}`);
    report.artifacts.push({ file, sha256 });
  }
  for (const [engine, launcher] of Object.entries({ chromium, webkit, firefox })) {
    const browser = await launcher.launch();
    try {
      const result = { engine, version: browser.version(), rounds: [] };
      report.results.push(result);
      save();
      for (let round = 0; round < 3; round++) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
        try {
          const page = await context.newPage();
          await page.goto(base + '#budget');
          await page.waitForFunction(() => !!window.opensheet?.selection.get());
          const samples = await page.evaluate(
            async (file) => {
              const app = window.opensheet;
              await app.load(file);
              let book = app.getWorkbook(),
                sheet = book.getSheets()[0];
              const range = { startRow: 0, endRow: 100, startColumn: 0, endColumn: 100 };
              const out = [];
              for (const scenario of ['numeric', 'datetime']) {
                const date = scenario === 'datetime';
                const matrix = Array.from({ length: 100 }, (_, r) =>
                  Array.from({ length: 100 }, (_, c) => {
                    const i = r * 100 + c;
                    return date ? 45293 + (i % 1440) / 1440 : i + 0.25;
                  }),
                );
                await sheet.range(range).setValues(matrix);
                await sheet
                  .range(range)
                  .setStyle({ numberFormat: date ? 'yyyy-mm-dd hh:mm' : '#,##0.00' });
                const fmt = new Intl.NumberFormat('en-US', {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                  useGrouping: true,
                });
                function verify(values) {
                  if (values.length !== 100) throw new Error('Row count mismatch');
                  for (let r = 0; r < 100; r++) {
                    if (values[r].length !== 100) throw new Error('Column count mismatch');
                    for (let c = 0; c < 100; c++) {
                      const i = r * 100 + c,
                        minute = i % 1440;
                      const expected = date
                        ? `2024-01-02 ${String(Math.floor(minute / 60)).padStart(2, '0')}:${String(minute % 60).padStart(2, '0')}`
                        : fmt.format(i + 0.25);
                      if (values[r][c] !== expected)
                        throw new Error(
                          `Mismatch ${scenario} ${r},${c}: ${values[r][c]} != ${expected}`,
                        );
                    }
                  }
                }
                const frames = [],
                  reads = [];
                let active = true,
                  previous = performance.now(),
                  frame;
                const tick = (now) => {
                  frames.push(now - previous);
                  previous = now;
                  if (active) frame = requestAnimationFrame(tick);
                };
                frame = requestAnimationFrame(tick);
                for (let i = 0; i < 5; i++) {
                  const start = performance.now();
                  const values = await sheet.range(range).getDisplayValues();
                  reads.push(performance.now() - start);
                  verify(values);
                }
                active = false;
                cancelAnimationFrame(frame);
                const id = book.id;
                await app.open(id);
                book = app.getWorkbook();
                sheet = book.getSheets()[0];
                verify(await sheet.range(range).getDisplayValues());
                out.push({ scenario, readsMs: reads, frameIntervalsMs: frames });
              }
              return out;
            },
            createWorkbookFile({ sheets: [{ name: 'Format workload', rows: 100, columns: 100 }] }),
          );
          result.rounds.push({ round, samples });
          save();
          process.stdout.write(
            JSON.stringify({
              engine,
              round,
              samples: samples.map(({ scenario, readsMs, frameIntervalsMs }) => ({
                scenario,
                readsMs,
                maxFrameMs: Math.max(0, ...frameIntervalsMs),
              })),
            }) + '\n',
          );
        } finally {
          await context.close();
        }
      }
    } finally {
      await browser.close();
    }
  }
  if (JSON.stringify(localBuildFingerprint()) !== JSON.stringify(report.build))
    throw new Error('Build changed during measurement');
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  save();
  throw error;
}

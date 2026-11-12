import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/recalculation-worker-timing.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const browser = await { chromium, firefox, webkit }[engine].launch();
const report = {
  startedAt: new Date().toISOString(),
  engine,
  browser: browser.version(),
  sourceFingerprint: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Development-server direct Workbook Worker with opt-in operation timing; three fresh contexts, five edits per context. All outputs checked after timing. Queue, Worker execution and round trip include instrumentation and scheduling, not CPU attribution. No production asset, UI, peak memory, other devices or reference acceptance.',
  rounds: [],
};
writeFileSync(output, JSON.stringify(report, null, 2), { flag: 'wx' });
try {
  for (let round = 0; round < 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5173/');
      const result = await page.evaluate(
        async ({ moduleUrl, file }) => {
          const { Workbook } = await import(/* @vite-ignore */ moduleUrl);
          let lastRead;
          const book = await new Workbook(file, {
            onOperationTiming: (timing) => {
              if (timing.method === 'read') lastRead = timing;
            },
          }).ready();
          try {
            const [source, fan, chain, overlap] = book.getSheets();
            await source.range('A1:A32').setValues(Array.from({ length: 32 }, () => [1]));
            await fan
              .range('A1:B128')
              .setFormulas(
                Array.from({ length: 128 }, (_, i) => [`Source!$A$1+${i}`, `${i}+1000`]),
              );
            await chain
              .range('A1:A128')
              .setFormulas(Array.from({ length: 128 }, (_, i) => [i ? `A${i}+1` : 'Source!$A$1']));
            await overlap
              .range('A1:A128')
              .setFormulas(Array.from({ length: 128 }, (_, i) => [`SUM(Source!$A$1:$A$32)+${i}`]));
            const verify = async (target) => {
              const outputs = [];
              for (const [name, sheet, range] of [
                ['fanout', fan, 'A1:B128'],
                ['chain', chain, 'A1:A128'],
                ['overlap', overlap, 'A1:A128'],
              ]) {
                lastRead = undefined;
                const start = performance.now();
                const values = await sheet.range(range).getValues();
                const elapsedMs = performance.now() - start;
                if (!lastRead) throw new Error('Missing Worker read timing');
                outputs.push({ name, elapsedMs, timing: lastRead });
                for (let i = 0; i < 128; i++) {
                  const expected = name === 'overlap' ? target + 31 + i : target + i;
                  if (values[i][0] !== expected || (name === 'fanout' && values[i][1] !== i + 1000))
                    throw new Error(`Wrong ${name} output ${i}`);
                }
              }
              return outputs;
            };
            await verify(1);
            const samples = [];
            for (let target = 2; target <= 6; target++) {
              await source.range('A1').setValues([[target]]);
              samples.push({ target, outputs: await verify(target) });
            }
            return { samples, verifiedOutputsPerSample: 512 };
          } finally {
            await book.deleteStorage();
            book.dispose();
          }
        },
        {
          moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
          file: createWorkbookFile({
            sheets: ['Source', 'Fan', 'Chain', 'Overlap'].map((name) => ({
              name,
              rows: 128,
              columns: 2,
            })),
          }),
        },
      );
      report.rounds.push({ round: round + 1, ...result });
    } catch (error) {
      report.rounds.push({ round: round + 1, error: String(error) });
      throw error;
    } finally {
      await context.close();
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    }
  }
} finally {
  await browser.close();
}

import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/delete-workbook-benchmark.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const base = 'http://127.0.0.1:5186/';
const report = {
  engine,
  startedAt: new Date().toISOString(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  artifacts: [],
  rounds: [],
  scope:
    'Local headless production preview; 1m generated cells and 256 synthetic scratch records. Timing covers deletion through acknowledgement, excludes fixture generation and verification. No isolated process/Worker peak memory, physical power-loss, hardware or competing-workload measurements.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const file of [
  'index.html',
  ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
]) {
  const response = await fetch(new URL(file, base));
  const digest = hash(new Uint8Array(await response.arrayBuffer()));
  if (!response.ok || digest !== hash(readFileSync(`apps/playground/dist/${file}`)))
    throw new Error(`Build mismatch: ${file}`);
  report.artifacts.push({ file, sha256: digest });
}
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  report.browser = browser.version();
  for (let round = 1; round <= 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto(base + '#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(async () => {
        const other = window.opensheet.getWorkbook();
        const otherValues = JSON.stringify(await other.getSheets()[0].range('A1:F11').getValues());
        const snapshot = await other.toJSON();
        const book = new other.constructor({
          ...snapshot,
          workbookId: `delete_bench_${crypto.randomUUID()}`,
        });
        await book.ready();
        await book.generate(1_000_000);
        const id = book.id;
        const bytes = (await book.storageStats()).bytes;
        const get = (request) =>
          new Promise((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          });
        const db = await get(indexedDB.open('opensheet'));
        const tx = db.transaction('scratch', 'readwrite');
        for (let i = 0; i < 256; i++)
          tx.objectStore('scratch').put(new Uint8Array(1024), `${id}/bench/${i}`);
        await new Promise((resolve, reject) => {
          tx.oncomplete = resolve;
          tx.onabort = () => reject(tx.error);
        });
        const counts = async (workbookId) => {
          const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
          const range = IDBKeyRange.bound(`${workbookId}/`, `${workbookId}/\uffff`);
          return Promise.all([
            get(tx.objectStore('heads').count(workbookId)),
            ...['records', 'sizes', 'scratch'].map((name) =>
              get(tx.objectStore(name).count(range)),
            ),
          ]);
        };
        const capture = async (workbookId) => {
          const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
          const range = IDBKeyRange.bound(`${workbookId}/`, `${workbookId}/\uffff`);
          const requests = [
            tx.objectStore('heads').get(workbookId),
            ...['records', 'sizes', 'scratch'].flatMap((name) => [
              tx.objectStore(name).getAllKeys(range),
              tx.objectStore(name).getAll(range),
            ]),
          ];
          return JSON.stringify(await Promise.all(requests.map(get)), (_key, value) =>
            value instanceof Uint8Array ? Array.from(value) : value,
          );
        };
        const before = await counts(id),
          otherBefore = await counts(other.id),
          otherRaw = await capture(other.id);
        if (before[0] !== 1 || !before[1] || before[1] !== before[2] || before[3] !== 256)
          throw new Error(`Unexpected fixture ${before}`);
        const start = performance.now();
        await book.deleteStorage();
        const deleteMs = performance.now() - start;
        const after = await counts(id),
          otherAfter = await counts(other.id);
        if (after.some(Boolean)) throw new Error(`Deletion left records: ${after}`);
        if (
          JSON.stringify(otherBefore) !== JSON.stringify(otherAfter) ||
          otherRaw !== (await capture(other.id)) ||
          JSON.stringify(await other.getSheets()[0].range('A1:F11').getValues()) !== otherValues
        )
          throw new Error('Other workbook changed');
        db.close();
        const otherId = other.id;
        await other.close();
        await window.opensheet.open(otherId);
        if (
          JSON.stringify(
            await window.opensheet.getWorkbook().getSheets()[0].range('A1:F11').getValues(),
          ) !== otherValues
        )
          throw new Error('Other workbook differs after reopen');
        return {
          cells: 1_000_000,
          bytes,
          before,
          after,
          otherBefore,
          otherAfter,
          otherRawUnchanged: true,
          otherReopened: true,
          deleteMs,
        };
      });
      report.rounds.push({ round, ...result });
      console.log(JSON.stringify({ engine, round, ...result }));
    } catch (error) {
      report.rounds.push({ round, error: String(error) });
      throw error;
    } finally {
      save();
      await context.close();
    }
  }
  report.finishedAt = new Date().toISOString();
} finally {
  save();
  await browser.close();
}

import { chromium, firefox, webkit } from '@playwright/test';
import { cpus, platform, release } from 'node:os';
import { writeFileSync } from 'node:fs';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium', count = '1000000'] = process.argv.slice(2);
const cells = Number(count);
if (
  !output ||
  !['chromium', 'firefox', 'webkit'].includes(engine) ||
  !Number.isSafeInteger(cells) ||
  cells < 1 ||
  cells > 10_000_000
)
  throw new Error(
    'Usage: node scripts/recovery-reopen-benchmark.mjs NEW_REPORT [chromium|firefox|webkit] [cells]',
  );
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  const report = {
    engine,
    browser: browser.version(),
    node: process.version,
    platform: platform(),
    release: release(),
    cpu: cpus()[0]?.model,
    source: sourceFingerprint(),
    build: localBuildFingerprint(),
    startedAt: new Date().toISOString(),
    rounds: [],
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  for (let round = 1; round <= 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5186/#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const data = await page.evaluate(
        async ({ dirtyFirst, cells }) => {
          const app = window.opensheet;
          const book = app.getWorkbook();
          await book.generate(cells);
          const id = book.id;
          const revision = book.revision;
          const baselineBytes = (await book.storageStats()).bytes;
          const sample = async (current) => await current.getSheets()[0].range('A1:B2').getValues();
          const values = JSON.stringify(await sample(book));
          const get = (request) =>
            new Promise((resolve, reject) => {
              request.onsuccess = () => resolve(request.result);
              request.onerror = () => reject(request.error);
            });
          const encoder = new TextEncoder();
          const txDone = (tx) =>
            new Promise((resolve, reject) => {
              tx.oncomplete = resolve;
              tx.onabort = () => reject(tx.error);
            });
          const orphans = 4096;
          const seed = async () => {
            const db = await get(indexedDB.open('opensheet'));
            try {
              const tx = db.transaction(['records', 'sizes', 'scratch'], 'readwrite');
              const done = txDone(tx);
              const records = tx.objectStore('records'),
                sizes = tx.objectStore('sizes');
              for (let i = 0; i < orphans; i++) {
                const key = `${id}/zz-orphan/${i}`;
                const page = new Uint8Array(1025);
                page[0] = 1;
                records.put(page, key);
                const ledger = encoder.encode(JSON.stringify([key, page.byteLength]));
                const encoded = new Uint8Array(ledger.length + 1);
                encoded.set(ledger, 1);
                sizes.put(encoded, key);
              }
              tx.objectStore('scratch').put(new Uint8Array([1]), `${id}/sort/abandoned`);
              await done;
            } finally {
              db.close();
            }
          };
          const reopen = async (dirty) => {
            await app.getWorkbook().close();
            if (dirty) await seed();
            const start = performance.now();
            await app.open(id);
            const ms = performance.now() - start;
            const current = app.getWorkbook();
            if (
              current.revision !== revision ||
              JSON.stringify(await sample(current)) !== values ||
              (await current.storageStats()).bytes !== baselineBytes
            )
              throw new Error('Recovered workbook differs from baseline');
            const db = await get(indexedDB.open('opensheet'));
            try {
              const tx = db.transaction(['records', 'sizes', 'scratch']);
              const range = IDBKeyRange.bound(`${id}/zz-orphan/`, `${id}/zz-orphan/\uffff`);
              const counts = await Promise.all([
                get(tx.objectStore('records').count(range)),
                get(tx.objectStore('sizes').count(range)),
                get(tx.objectStore('scratch').count(IDBKeyRange.bound(`${id}/`, `${id}/\uffff`))),
              ]);
              if (counts.some(Boolean)) throw new Error('Recovery left staging records');
            } finally {
              db.close();
            }
            return ms;
          };
          const timings = {};
          for (const dirty of dirtyFirst ? [true, false] : [false, true]) {
            timings[dirty ? 'dirtyMs' : 'cleanMs'] = await reopen(dirty);
          }
          return { baselineBytes, orphans, cells, ...timings };
        },
        { dirtyFirst: round % 2 === 0, cells },
      );
      report.rounds.push({ round, ...data });
      console.log(JSON.stringify({ round, ...data }));
    } catch (error) {
      report.rounds.push({ round, error: String(error) });
      throw error;
    } finally {
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
      await context.close();
    }
  }
} finally {
  await browser.close();
}

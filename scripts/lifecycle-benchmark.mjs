import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { cpus, totalmem } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';
import { startProcessMemorySampler } from './process-memory-sampler.mjs';
const option = (name, fallback) =>
  process.argv.find((arg) => arg.startsWith(`--${name}=`))?.slice(name.length + 3) ?? fallback;
const output = option('output', ''),
  cycles = Number(option('cycles', '100')),
  fixture = option('fixture', 'scalar');
if (!['scalar', 'metadata', 'history'].includes(fixture)) throw new Error('Unknown fixture');
if (!output || existsSync(output) || !Number.isSafeInteger(cycles) || cycles < 1 || cycles > 1000)
  throw new Error('Use a new --output path and 1–1000 cycles');
const launchers = { chromium, webkit, firefox },
  engines = option('engines', 'chromium,webkit,firefox').split(',');
if (engines.some((name) => !launchers[name])) throw new Error('Unknown engine');
const processMemory = process.argv.includes('--process-memory');
if (processMemory && engines.some((name) => name !== 'chromium'))
  throw new Error('Process memory sampling currently supports Chromium only');
const report = {
  startedAt: new Date().toISOString(),
  cycles,
  fixture,
  url: option('url', 'http://127.0.0.1:5186/'),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  host: { node: process.version, cpu: cpus()[0].model, memoryBytes: totalmem() },
  scope:
    'Fresh isolated browser contexts. Each cycle creates, edits, reattaches twice, replaces, reopens and disposes an editor; owned storage is deleted. Tracks newly constructed Workers/BroadcastChannels/ResizeObservers and workbook data/commit/dispose subscriptions. Counts measure explicit terminate/close/disconnect/unsubscribe calls, not physical process retention. Chromium main-target heap/DOM samples follow requested GC every 10 cycles; Worker heap is unavailable. ' +
    (processMemory
      ? 'Browser process RSS is sampled and checkpointed after requested main-target GC; shared pages can be counted repeatedly and Worker memory is not isolated. '
      : 'Browser process memory is unavailable. ') +
    'No hard heap threshold or universal leak-free certification.',
  results: [],
};
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
try {
  const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
  report.artifacts = [];
  for (const file of [
    'index.html',
    ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
  ]) {
    const response = await fetch(new URL(file, report.url));
    const sha256 = hash(new Uint8Array(await response.arrayBuffer()));
    if (!response.ok || sha256 !== hash(readFileSync(`apps/playground/dist/${file}`)))
      throw new Error(`Build mismatch: ${file}`);
    report.artifacts.push({ file, sha256 });
  }
  for (const engine of engines) {
    const browser = await launchers[engine].launch();
    const memorySampler = processMemory ? startProcessMemorySampler() : undefined;
    let result;
    try {
      const context = await browser.newContext(),
        page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.stack ?? String(error)));
      await page.goto(report.url);
      await page.waitForFunction(
        () =>
          !!window.opensheet?.selection.get() && !document.querySelector('#scene-select').disabled,
      );
      await page.evaluate(async (fixture) => {
        await window.opensheet.getGrid().ready();
        const workers = new Set(),
          channels = new Set(),
          observers = new Set(),
          subscriptions = new Set();
        const NativeWorker = Worker,
          NativeChannel = BroadcastChannel,
          NativeObserver = ResizeObserver;
        window.Worker = class extends NativeWorker {
          constructor(...args) {
            super(...args);
            workers.add(this);
          }
          terminate() {
            workers.delete(this);
            super.terminate();
          }
        };
        window.BroadcastChannel = class extends NativeChannel {
          constructor(...args) {
            super(...args);
            channels.add(this);
          }
          close() {
            channels.delete(this);
            super.close();
          }
        };
        window.ResizeObserver = class extends NativeObserver {
          observe(...args) {
            observers.add(this);
            return super.observe(...args);
          }
          disconnect() {
            observers.delete(this);
            super.disconnect();
          }
        };
        const prototype = Object.getPrototypeOf(window.opensheet.getWorkbook());
        for (const method of ['onData', 'onCommit', 'onDispose']) {
          const original = prototype[method];
          prototype[method] = function (...args) {
            const token = {};
            subscriptions.add(token);
            const stop = original.apply(this, args);
            return () => {
              subscriptions.delete(token);
              stop();
            };
          };
        }
        window.__lifecycleCycle = async (cycle) => {
          const host = document.createElement('div');
          host.style.height = '300px';
          document.body.append(host);
          const Editor = window.opensheet.constructor,
            editor = new Editor({ container: host });
          const started = performance.now();
          try {
            let book = await editor.createWorkbook({
              sheets: [{ name: 'Cycle', rows: 100, columns: 8 }],
            });
            await book
              .getSheets()[0]
              .range('A1')
              .setValues([[cycle]]);
            if (fixture === 'metadata' || fixture === 'history') {
              await book.transaction({}, async (tx) => {
                const sheet = tx.getSheets()[0];
                await sheet
                  .range('B1:B16')
                  .setValues(Array.from({ length: 16 }, (_, i) => [i + cycle]));
                await sheet
                  .range('E1:E16')
                  .setFormulas(Array.from({ length: 16 }, (_, i) => [`B${i + 1}*2`]));
                for (let row = 1; row <= 16; row++) {
                  await sheet.range(`A${row}`).setStyle({
                    bold: row % 2 === 0,
                    color: `#${row.toString(16).padStart(6, '0')}`,
                  });
                  await sheet.range(`C${row}:D${row}`).merge();
                }
              });
            }
            await editor.attachWorkbook(book);
            await editor.attachWorkbook(book);
            if (fixture === 'history') {
              for (let value = 1; value <= 20; value++)
                await book
                  .getSheets()[0]
                  .range('H1')
                  .setValues([[value]]);
              for (let i = 0; i < 5; i++) await book.undo();
              if ((await book.getSheets()[0].range('H1').getValues())[0][0] !== 15)
                throw new Error('Lifecycle undo mismatch');
              await editor.open(book.id);
              book = editor.getWorkbook();
              if ((await book.getSheets()[0].range('H1').getValues())[0][0] !== 15)
                throw new Error('Lifecycle history reopen mismatch');
              for (let i = 0; i < 5; i++) await book.redo();
              if ((await book.getSheets()[0].range('H1').getValues())[0][0] !== 20)
                throw new Error('Lifecycle redo mismatch');
            }
            const file = await book.toJSON();
            await book.deleteStorage();
            if (fixture === 'history')
              await new Promise((resolve) => requestAnimationFrame(() => resolve()));
            await editor.load(file);
            book = editor.getWorkbook();
            await editor.open(book.id);
            book = editor.getWorkbook();
            if ((await book.getSheets()[0].range('A1').getValues())[0][0] !== cycle)
              throw new Error('Lifecycle data mismatch');
            if (fixture === 'metadata' || fixture === 'history') {
              const sheet = book.getSheets()[0];
              const calculated = await sheet.range('E1:E16').getValues();
              const formatted = await book.readCells(
                sheet.id,
                { startRow: 0, endRow: 16, startColumn: 0, endColumn: 1 },
                { fields: ['style'] },
              );
              if (
                sheet.workbook.sheetData(sheet.id).merges.length !== 16 ||
                sheet.workbook
                  .sheetData(sheet.id)
                  .merges.some(
                    (merge, i) =>
                      merge.startRow !== i ||
                      merge.endRow !== i + 1 ||
                      merge.startColumn !== 2 ||
                      merge.endColumn !== 4,
                  ) ||
                calculated.length !== 16 ||
                formatted.cells.length !== 16 ||
                calculated.some((row, i) => row[0] !== (i + cycle) * 2) ||
                formatted.cells.some(
                  (cell, i) => cell.style?.color !== `#${(i + 1).toString(16).padStart(6, '0')}`,
                )
              )
                throw new Error('Lifecycle metadata mismatch');
            }
            if (
              fixture === 'history' &&
              (await book.getSheets()[0].range('H1').getValues())[0][0] !== 20
            )
              throw new Error('Lifecycle replacement history value mismatch');
            await book.deleteStorage();
          } finally {
            editor.dispose();
            host.remove();
          }
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          const counts = {
            workers: workers.size,
            channels: channels.size,
            observers: observers.size,
            subscriptions: subscriptions.size,
          };
          if (Object.values(counts).some((count) => count !== 0))
            throw new Error(`Retained resources: ${JSON.stringify(counts)}`);
          return { cycle, durationMs: performance.now() - started, counts };
        };
      }, fixture);
      result = { engine, browserVersion: browser.version(), cycles: [], memorySamples: [] };
      if (memorySampler) result.processMemoryCheckpoints = [];
      report.results.push(result);
      save();
      const cdp = engine === 'chromium' ? await context.newCDPSession(page) : null;
      const sample = async (cycle) => {
        if (!cdp) {
          result.memorySamples.push({
            cycle,
            unavailable: 'No CDP main-target GC/heap measurement for this engine',
          });
          return;
        }
        await cdp.send('HeapProfiler.collectGarbage');
        const instances = {};
        for (const [name, expression] of Object.entries({
          editor: 'Object.getPrototypeOf(window.opensheet)',
          workbook: 'Object.getPrototypeOf(window.opensheet.getWorkbook())',
          grid: 'Object.getPrototypeOf(window.opensheet.getGrid())',
        })) {
          const prototype = await cdp.send('Runtime.evaluate', { expression });
          let objects;
          try {
            objects = await cdp.send('Runtime.queryObjects', {
              prototypeObjectId: prototype.result.objectId,
            });
            const count = await cdp.send('Runtime.callFunctionOn', {
              objectId: objects.objects.objectId,
              functionDeclaration: 'function() { return this.length; }',
              returnByValue: true,
            });
            instances[name] = count.result.value;
          } finally {
            if (objects?.objects.objectId)
              await cdp.send('Runtime.releaseObject', { objectId: objects.objects.objectId });
            await cdp.send('Runtime.releaseObject', { objectId: prototype.result.objectId });
          }
        }
        result.memorySamples.push({
          cycle,
          instances,
          heap: await cdp.send('Runtime.getHeapUsage'),
          dom: await cdp.send('Memory.getDOMCounters'),
        });
        if (memorySampler) {
          // Drain a periodic sample that may have started before the GC checkpoint.
          await memorySampler.take();
          const count = memorySampler.snapshot().samples.length;
          await memorySampler.take();
          const samples = memorySampler.snapshot().samples;
          result.processMemoryCheckpoints.push({
            cycle,
            sample: samples.length > count ? samples.at(-1) : null,
          });
        }
      };
      await sample(0);
      for (let cycle = 1; cycle <= cycles; cycle++) {
        result.cycles.push(await page.evaluate((cycle) => window.__lifecycleCycle(cycle), cycle));
        if (cycle % 10 === 0 || cycle === cycles) {
          await sample(cycle);
          save();
          process.stdout.write(
            JSON.stringify({
              engine,
              cycle,
              counts: result.cycles.at(-1).counts,
              memory: result.memorySamples.at(-1),
            }) + '\n',
          );
        }
      }
      if (errors.length) throw new Error(errors.join('\n'));
      result.databaseCleanup = await page.evaluate(async () => {
        window.opensheet.dispose();
        return new Promise((resolve, reject) => {
          const req = indexedDB.deleteDatabase('opensheet');
          const timer = setTimeout(() => reject(new Error('Database cleanup timed out')), 5000);
          req.onsuccess = () => {
            clearTimeout(timer);
            resolve('deleted');
          };
          req.onerror = () => {
            clearTimeout(timer);
            reject(req.error);
          };
        });
      });
      save();
      await context.close();
    } finally {
      if (memorySampler) {
        await memorySampler.stop();
        if (result) result.processMemory = memorySampler.snapshot();
        save();
      }
      await browser.close();
    }
  }
  if (JSON.stringify(localBuildFingerprint()) !== JSON.stringify(report.build))
    throw new Error('Build changed during run');
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  save();
  throw error;
}

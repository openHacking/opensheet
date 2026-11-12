import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';
import { existsSync, writeFileSync } from 'node:fs';
const url = process.argv.find((v) => v.startsWith('--url='))?.slice(6) ?? 'http://127.0.0.1:4174/';
const outputPath = process.argv.find((v) => v.startsWith('--output='))?.slice(9);
const engine = process.argv.find((v) => v.startsWith('--browser='))?.slice(10) ?? 'chromium';
if (!outputPath || existsSync(outputPath)) throw new Error('Supply a new --output path');
if (!['chromium', 'firefox', 'webkit'].includes(engine)) throw new Error('Invalid browser');
const count = Number(process.argv.find((v) => v.startsWith('--cells='))?.slice(8) ?? 1000000);
if (!Number.isSafeInteger(count) || count < 10000 || count > 10000000 || count % 10000)
  throw new Error('Invalid size');
const source = sourceFingerprint(),
  build = localBuildFingerprint();
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  const page = await browser.newPage();
  await page.goto(url + '#performance');
  await page.waitForFunction(
    () =>
      document.querySelector('#scene-title')?.textContent === 'Performance Lab' &&
      !document.querySelector('#scene-select')?.disabled &&
      !!window.opensheet,
  );
  const result = await page.evaluate(async (count) => {
    const app = window.opensheet,
      book = app.getWorkbook();
    await book.generate(count);
    const started = performance.now(),
      file = await book.toJSON(),
      toJSONMs = performance.now() - started;
    let expectedCount = 0,
      expectedHash = 0;
    const hash = (row, column, input) => {
      let h = 2166136261;
      for (const char of `${row}:${column}:${JSON.stringify(input)}`)
        h = Math.imul(h ^ char.charCodeAt(0), 16777619);
      return h >>> 0;
    };
    const kinds = { numbers: 'number', strings: 'string', booleans: 'boolean' };
    for (const block of file.sheets[0].blocks)
      for (const [name, type] of Object.entries(kinds)) {
        const col = block.data[name];
        if (!col) continue;
        col.positions.forEach((offset, i) => {
          expectedCount++;
          expectedHash =
            (expectedHash +
              hash(block.row * 64 + Math.floor(offset / 32), block.column * 32 + (offset % 32), {
                type,
                value: col.values[i],
              })) >>>
            0;
        });
      }
    if (expectedCount !== count) throw new Error('Complete JSON omitted cells');
    const json = JSON.stringify(file),
      jsonBytes = new TextEncoder().encode(json).length;
    let streamJSONBytes = 0;
    for await (const text of book.streamJSON())
      streamJSONBytes += new TextEncoder().encode(text).length;
    if (streamJSONBytes !== jsonBytes) throw new Error('Streaming JSON size differs');
    const verify = async () => {
      let cursor,
        seen = 0,
        actualHash = 0;
      do {
        const batch = await book.scanCells(book.getSheets()[0].id, {
          cursor,
          fields: ['input'],
          limit: 2048,
        });
        for (const cell of batch.cells) {
          seen++;
          actualHash = (actualHash + hash(cell.row, cell.column, cell.input)) >>> 0;
        }
        cursor = batch.cursor;
      } while (cursor);
      if (seen !== count || actualHash !== expectedHash) throw new Error('Imported data differs');
    };
    const measureImport = async (action) => {
      const sent = performance.now(),
        phases = {};
      const unsubscribe = book.onProgress((event) => {
        const time = performance.now() - sent;
        const phase = (phases[event.phase] ??= { firstMs: time, events: 0 });
        phase.lastMs = time;
        phase.events++;
        phase.completed = event.completed;
        phase.total = event.total;
        if (event.elapsedMs !== undefined) phase.elapsedMs = event.elapsedMs;
      });
      try {
        await action();
        return { roundTripMs: performance.now() - sent, phases };
      } finally {
        unsubscribe();
      }
    };
    let time = performance.now();
    const jsonProfile = await measureImport(() =>
      book.importJSON(new Blob([json], { type: 'application/json' })),
    );
    const importJSONMs = performance.now() - time;
    await verify();
    time = performance.now();
    const parts = [];
    for await (const part of book.streamBinary()) parts.push(part);
    const binary = new Blob(parts),
      streamBinaryMs = performance.now() - time;
    time = performance.now();
    const binaryProfile = await measureImport(() => book.importBinary(binary));
    const importBinaryMs = performance.now() - time;
    await verify();
    return {
      cells: count,
      toJSONMs,
      jsonBytes,
      streamJSONBytes,
      importJSONMs,
      jsonProfile,
      binaryBytes: binary.size,
      streamBinaryMs,
      importBinaryMs,
      binaryProfile,
      verifiedCellsAfterEachImport: count,
      verificationHash: expectedHash,
      storage: await book.storageStats(),
    };
  }, count);
  const output = {
    date: new Date().toISOString(),
    browser: browser.version(),
    engine,
    source,
    build,
    url,
    profile:
      'Progress instrumentation changes timing; event windows use main-page performance.now, phase elapsedMs comes from the Worker. Unaccounted time includes parsing, staging and publication.',
    fixture:
      'deterministic 100-column mixed workload; complete JSON materialization and JSON Blob/binary Blob round trips; count plus all-cell order-independent input checksum',
    result,
  };
  writeFileSync(outputPath, JSON.stringify(output, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify(output, null, 2));
} finally {
  await browser.close();
}

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
    const app = window.opensheet;
    let book = app.getWorkbook();
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
    // Independent reconstruction of the fixed-seed workload, without using exported cell values.
    const expected = (n) => {
      let state = (Math.imul(n + 1, 1664525) + 1013904223) >>> 0;
      if (n % 10 < 6) return n === 0 ? 0 : state / 1000;
      if (n % 10 === 9) return Math.floor(n / 10) % 2 === 1;
      let text = '';
      for (let i = 0; i < 32; i++) {
        state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
        text += String.fromCharCode(33 + (state % 94));
      }
      return text;
    };
    const verify = async () => {
      let cursor,
        seen = 0;
      const visited = new Uint8Array(count);
      do {
        const batch = await book.scanCells(book.getSheets()[0].id, {
          cursor,
          fields: ['input'],
          limit: 2048,
        });
        for (const cell of batch.cells) {
          const index = cell.row * 100 + cell.column;
          if (
            !Number.isInteger(index) ||
            index < 0 ||
            index >= count ||
            cell.column < 0 ||
            cell.column >= 100 ||
            visited[index]
          )
            throw new Error(`Unexpected or duplicate cell ${cell.row}:${cell.column}`);
          visited[index] = 1;
          const value = expected(index);
          if (cell.input?.type !== typeof value || !Object.is(cell.input.value, value))
            throw new Error(`Cell differs at ${cell.row}:${cell.column}`);
          seen++;
        }
        cursor = batch.cursor;
      } while (cursor);
      if (seen !== count) throw new Error(`Missing cells: ${seen}/${count}`);
    };
    const reopenAndVerify = async () => {
      const id = book.id;
      await app.open(id, {
        database: 'opensheet-performance',
        preserveOrphans: true,
        budgetBytes: 256 * 1024 * 1024,
      });
      book = app.getWorkbook();
      await verify();
    };
    await verify();
    let time = performance.now();
    await book.importJSON(new Blob([json], { type: 'application/json' }));
    const importJSONMs = performance.now() - time;
    await verify();
    await reopenAndVerify();
    time = performance.now();
    const parts = [];
    for await (const part of book.streamBinary()) parts.push(part);
    const binary = new Blob(parts),
      streamBinaryMs = performance.now() - time;
    time = performance.now();
    await book.importBinary(binary);
    const importBinaryMs = performance.now() - time;
    await verify();
    await reopenAndVerify();
    return {
      cells: count,
      toJSONMs,
      jsonBytes,
      streamJSONBytes,
      importJSONMs,
      binaryBytes: binary.size,
      streamBinaryMs,
      importBinaryMs,
      exactVerifiedCellsPerPass: count,
      exactVerificationPasses: 5,
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
    fixture:
      'deterministic 100-column mixed workload; complete JSON materialization and JSON Blob/binary Blob round trips; independent per-cell value, type, coordinate and uniqueness checks before import and before/after reopening each import; hash retained only as extra diagnostic',
    result,
  };
  writeFileSync(outputPath, JSON.stringify(output, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify(output, null, 2));
} finally {
  await browser.close();
}

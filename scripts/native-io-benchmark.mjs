import { chromium } from '@playwright/test';
import { writeFileSync } from 'node:fs';
const url = process.argv.find((v) => v.startsWith('--url='))?.slice(6) ?? 'http://127.0.0.1:4174/';
const count = Number(process.argv.find((v) => v.startsWith('--cells='))?.slice(8) ?? 1000000);
if (!Number.isSafeInteger(count) || count < 10000 || count > 10000000 || count % 10000)
  throw new Error('Invalid size');
const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  await page.goto(url + '#performance');
  await page.waitForFunction(() => !document.querySelector('#scene-select')?.disabled);
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
    let time = performance.now();
    await book.importJSON(new Blob([json], { type: 'application/json' }));
    const importJSONMs = performance.now() - time;
    await verify();
    time = performance.now();
    const parts = [];
    for await (const part of book.streamBinary()) parts.push(part);
    const binary = new Blob(parts),
      streamBinaryMs = performance.now() - time;
    time = performance.now();
    await book.importBinary(binary);
    const importBinaryMs = performance.now() - time;
    await verify();
    return {
      cells: count,
      toJSONMs,
      jsonBytes,
      streamJSONBytes,
      importJSONMs,
      binaryBytes: binary.size,
      streamBinaryMs,
      importBinaryMs,
      verifiedCellsAfterEachImport: count,
      verificationHash: expectedHash,
      storage: await book.storageStats(),
    };
  }, count);
  const output = {
    date: new Date().toISOString(),
    browser: browser.version(),
    url,
    fixture:
      'deterministic 100-column mixed workload; complete JSON materialization and JSON Blob/binary Blob round trips; count plus all-cell order-independent input checksum',
    result,
  };
  writeFileSync('benchmarks/native-io.json', JSON.stringify(output, null, 2) + '\n');
  console.log(JSON.stringify(output, null, 2));
} finally {
  await browser.close();
}

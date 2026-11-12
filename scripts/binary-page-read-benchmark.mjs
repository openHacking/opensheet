import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Pass a new report path');
const path = resolve('packages/core/src/binary-file.ts');
const source = readFileSync(path, 'utf8');
const start = source.indexOf('    // Avoid an extra asynchronous Blob read for small pages');
const end = source.indexOf('\n  }\n}', start);
if (start < 0 || end < 0) throw new Error('Cannot reconstruct previous page reader');
const previous =
  source.slice(0, start) +
  `    const bytes = new Uint8Array(await this.source.slice(entry.offset, entry.offset + entry.length).arrayBuffer()), view = new DataView(bytes.buffer);
    assert(view.getUint32(0, true) === entry.sheet && view.getUint32(4, true) === entry.row && view.getUint32(8, true) === entry.column && view.getUint32(16, true) === bytes.length - 28 && crc32(bytes.subarray(28)) === view.getUint32(24, true), 'INVALID_ARGUMENT', 'Invalid page frame');
    return unpackPage(bytes.subarray(28), view.getUint32(12, true), view.getUint32(20, true));` +
  source.slice(end);
const scripts = {},
  hashes = {};
for (const variant of ['previous', 'current']) {
  const bundle = await build({
    stdin: {
      contents: `export {BinaryReader,writeBinary} from './packages/core/src/binary-file.ts'; export {encodeBlock,decodeBlock} from './packages/core/src/block-codec.ts'; export {createWorkbookFile} from './packages/core/src/file-codec.ts';`,
      resolveDir: process.cwd(),
    },
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: `Page_${variant}`,
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /binary-file\.ts$/ }, () => ({
            contents: variant === 'previous' ? previous : source,
            loader: 'ts',
            resolveDir: resolve('packages/core/src'),
          }));
        },
      },
    ],
  });
  scripts[variant] = bundle.outputFiles[0].text;
  hashes[variant] = createHash('sha256').update(scripts[variant]).digest('hex');
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hashes,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless browser, warm Blob, 8 real encoded workbook pages of 2048 pseudorandom text cells each; paired sequential reads excluding open, validation of every decoded value outside timing. Previous is reconstructed, not a clean historical build. No Worker, IndexedDB, cold I/O, peak memory, production assets or other devices.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), fixture: null, rounds: [] };
  report.results.push(result);
  try {
    const page = await browser.newPage();
    await page.route('http://127.0.0.1/**', (route) =>
      route.fulfill({ status: 200, contentType: 'text/html', body: '<!doctype html>' }),
    );
    await page.goto('http://127.0.0.1/');
    for (const script of Object.values(scripts)) await page.addScriptTag({ content: script });
    const measurement = await page.evaluate(async () => {
      const current = globalThis.Page_current;
      const metadata = current.createWorkbookFile({
        sheets: [{ name: 'Probe', rows: 512, columns: 32 }],
      });
      const expected = [],
        raw = [];
      let seed = 123456789;
      for (let p = 0; p < 8; p++) {
        const cells = {};
        for (let i = 0; i < 2048; i++) {
          let value = '';
          for (let j = 0; j < 128; j++) {
            seed ^= seed << 13;
            seed ^= seed >>> 17;
            seed ^= seed << 5;
            value += 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_'[
              (seed >>> 0) & 63
            ];
          }
          const row = p * 64 + Math.floor(i / 32),
            column = i % 32;
          cells[`${row}:${column}`] = {
            rowId: row,
            columnId: column,
            input: { type: 'string', value },
          };
        }
        expected.push(cells);
        raw.push(current.encodeBlock(cells, [], [], {}));
      }
      async function* pages() {
        for (let p = 0; p < raw.length; p++) yield { sheet: 0, row: p, column: 0, bytes: raw[p] };
      }
      const parts = [];
      for await (const part of current.writeBinary(metadata, pages())) parts.push(part);
      const blob = new Blob(parts),
        readers = {};
      for (const name of ['previous', 'current'])
        readers[name] = await globalThis[`Page_${name}`].BinaryReader.open(blob);
      const lengths = readers.current.directory.map((item) => item.length);
      if (lengths.some((length) => length <= 64 * 1024))
        throw new Error('Fixture misses large page path');
      const run = async (name) => {
        const start = performance.now(),
          result = [];
        for (let p = 0; p < raw.length; p++) result.push(await readers[name].page(p));
        const ms = performance.now() - start;
        for (let p = 0; p < raw.length; p++) {
          const decoded = current.decodeBlock(result[p]);
          const cells = [...decoded.records()];
          if (cells.length !== 2048) throw new Error('Wrong cell count');
          for (const cell of cells)
            if (cell.input.value !== expected[p][`${cell.rowId}:${cell.columnId}`]?.input.value)
              throw new Error('Wrong cell value');
        }
        return ms;
      };
      await run('previous');
      await run('current');
      const rounds = [];
      for (let i = 0; i < 8; i++) {
        const pair = {};
        for (const name of i % 2 ? ['current', 'previous'] : ['previous', 'current'])
          pair[name] = await run(name);
        rounds.push(pair);
      }
      return {
        fixture: { fileBytes: blob.size, rawBytes: raw.map((x) => x.length), frameBytes: lengths },
        rounds,
      };
    });
    Object.assign(result, measurement);
    save();
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();

import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { performance, PerformanceObserver } from 'node:perf_hooks';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
const diagnose = process.argv.includes('--diagnose');
const rows = Number(process.argv.find((arg) => arg.startsWith('--rows='))?.slice(7) ?? 100_000);
if (!output) throw new Error('Pass a new report path after building the packages');
if (!Number.isInteger(rows) || rows < 1 || rows > 100_000)
  throw new Error('Rows must be 1..100000');
const file = resolve('packages/core/src/binary-file.ts');
const source = readFileSync(file, 'utf8');
const start = source.indexOf("        const listed = parseBinaryJSON(directory, 'directory');");
const end = source.indexOf('        await reader.end();', start);
if (start < 0 || end < 0) throw new Error('Cannot reconstruct previous reader');
let previous =
  source.slice(0, start) +
  `        assert(JSON.stringify(parseBinaryJSON(directory, 'directory')) === JSON.stringify(entries), 'INVALID_ARGUMENT', 'Directory differs from pages');\n` +
  source.slice(end);
previous = previous.replace(
  'hasDirectoryFields(entry) &&',
  "entry !== null && typeof entry === 'object' && !Array.isArray(entry) &&",
);
previous = previous.replace(
  "    assert(hasDirectoryFields(entry), 'INVALID_ARGUMENT', 'Invalid directory entry');\n",
  '',
);
if (previous.includes('hasDirectoryFields(entry)'))
  throw new Error('Previous random reader reconstruction is incomplete');
const urls = [],
  hashes = [];
const marker = '    decodeFileMetadata(metadata);';
if (diagnose && source.split(marker).length !== 2)
  throw new Error('Cannot reconstruct metadata diagnostic');
const variants = diagnose
  ? [
      previous,
      source,
      source.replace(marker, '    // Diagnostic only: skip metadata decoding and validation.'),
    ]
  : [previous, source];
for (const contents of variants) {
  const built = await build({
    entryPoints: [file],
    bundle: true,
    platform: 'node',
    format: 'esm',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /binary-file\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/core/src'),
          }));
        },
      },
    ],
  });
  const text = built.outputFiles[0].text;
  hashes.push(createHash('sha256').update(text).digest('hex'));
  urls.push(`data:text/javascript;base64,${Buffer.from(text).toString('base64')}`);
}
const modules = await Promise.all(urls.map((url) => import(url)));
const count = 10_000;
const metadata = createWorkbookFile({ sheets: [{ name: 'Data', rows, columns: 1000 }] });
const headerText = new TextEncoder().encode(JSON.stringify(metadata));
const header = new Uint8Array(12),
  hv = new DataView(header.buffer);
header.set([79, 83, 66, 3]);
hv.setUint32(4, headerText.length, true);
hv.setUint32(8, modules[0].crc32(headerText), true);
const frame = new Uint8Array(29),
  fv = new DataView(frame.buffer);
frame[28] = 42;
fv.setUint32(12, 0, true);
fv.setUint32(16, 1, true);
fv.setUint32(20, 1, true);
fv.setUint32(24, modules[0].crc32(frame.subarray(28)), true);
const chunks = [header, headerText],
  entries = [];
let offset = header.length + headerText.length;
for (let i = 0; i < count; i++) {
  const row = Math.floor(i / 32),
    column = i % 32,
    data = frame.slice(),
    view = new DataView(data.buffer);
  view.setUint32(0, 0, true);
  view.setUint32(4, row, true);
  view.setUint32(8, column, true);
  chunks.push(data);
  entries.push({ sheet: 0, row, column, offset, length: data.length });
  offset += data.length;
}
const directory = new TextEncoder().encode(JSON.stringify(entries));
const directoryMarker = new Uint8Array(8),
  mv = new DataView(directoryMarker.buffer);
mv.setUint32(0, 0xffffffff, true);
mv.setUint32(4, directory.length, true);
const footer = new Uint8Array(24),
  foot = new DataView(footer.buffer);
footer.set([79, 83, 66, 69]);
foot.setBigUint64(4, BigInt(offset + 8), true);
foot.setBigUint64(12, BigInt(directory.length), true);
foot.setUint32(20, modules[0].crc32(directory), true);
const input = new Uint8Array(
  await new Blob([...chunks, directoryMarker, directory, footer]).arrayBuffer(),
);
const report = {
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  hashes,
  count,
  rows,
  bytes: input.length,
  scope: diagnose
    ? 'Diagnostic only: current reader versus a deliberately unsafe variant that skips metadata decoding/validation. Synthetic one-byte page payloads; not a valid workbook, production change, browser measurement or peak-memory measurement.'
    : 'Synthetic frame-level reader comparison; page payload is one byte, not a valid workbook page. Excludes IndexedDB, page decoding, browser memory and real data.',
  results: [],
};
const gcEvents = [];
const observer = diagnose
  ? new PerformanceObserver((list) => {
      for (const event of list.getEntries())
        gcEvents.push({
          start: event.startTime,
          duration: event.duration,
          kind: event.detail?.kind,
        });
    })
  : undefined;
observer?.observe({ entryTypes: ['gc'] });
for (const mode of diagnose ? ['metadata-diagnostic'] : ['open', 'stream']) {
  const run = async (which) => {
    const memoryBefore = diagnose ? process.memoryUsage() : undefined;
    const start = performance.now();
    if (mode === 'open' || mode === 'metadata-diagnostic') {
      const reader = await modules[which].BinaryReader.open(input);
      if (reader.directory.length !== count || (await reader.page(count - 1))[0] !== 42)
        throw new Error('Random reader mismatch');
    } else {
      let pages = 0;
      for await (const part of modules[which].readBinary(input))
        if ('bytes' in part) {
          if (part.bytes[0] !== 42) throw new Error('Page mismatch');
          pages++;
        }
      if (pages !== count) throw new Error('Stream page count mismatch');
    }
    const end = performance.now();
    if (!diagnose) return end - start;
    const memoryAfter = process.memoryUsage();
    await new Promise((resolve) => setImmediate(resolve));
    return {
      ms: end - start,
      start,
      end,
      heapBefore: memoryBefore.heapUsed,
      heapAfter: memoryAfter.heapUsed,
      rssBefore: memoryBefore.rss,
      rssAfter: memoryAfter.rss,
    };
  };
  const indexes = diagnose ? [1, 2] : [0, 1];
  await run(indexes[0]);
  await run(indexes[1]);
  const pairs = [];
  for (let i = 0; i < (diagnose ? 12 : 6); i++) {
    const times = [];
    for (const which of i % 2 ? [...indexes].reverse() : indexes) times[which] = await run(which);
    pairs.push(
      diagnose
        ? { current: times[1], withoutMetadataValidation: times[2] }
        : { beforeMs: times[0], afterMs: times[1] },
    );
  }
  report.results.push({ mode, pairs });
}
if (diagnose) {
  await new Promise((resolve) => setImmediate(resolve));
  observer.disconnect();
  report.gcEvents = gcEvents;
  for (const { pairs } of report.results)
    for (const pair of pairs)
      for (const sample of Object.values(pair))
        sample.gcOverlapMs = gcEvents.reduce(
          (sum, event) =>
            sum +
            Math.max(
              0,
              Math.min(sample.end, event.start + event.duration) -
                Math.max(sample.start, event.start),
            ),
          0,
        );
}
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });

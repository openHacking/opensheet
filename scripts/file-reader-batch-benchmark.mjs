import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { arch, cpus, platform, totalmem } from 'node:os';
import { resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const script = fileURLToPath(import.meta.url);
const mode = process.argv[2];
if (mode === '--sample') {
  const variant = process.argv[3];
  const { createWorkbookFile, FileReader } = await import(pathToFileURL(resolve(process.argv[4])));
  assert.ok(variant === 'matrix' || variant === 'batches');
  assert.equal(typeof global.gc, 'function');
  const file = createWorkbookFile({ sheets: [{ name: 'Sparse', rows: 1000, columns: 1000 }] });
  const reader = new FileReader(file);
  try {
    const range = reader.getSheetById(file.sheets[0].id).range('A1:ALL1000');
    global.gc();
    const before = process.memoryUsage();
    let cells = 0,
      maxBatch = 0,
      observedHeap = before.heapUsed;
    const check = (matrix) => {
      maxBatch = Math.max(maxBatch, matrix.length * 1000);
      for (const row of matrix) {
        assert.equal(row.length, 1000);
        for (const value of row) {
          assert.equal(value, null);
          cells++;
        }
      }
      observedHeap = Math.max(observedHeap, process.memoryUsage().heapUsed);
    };
    const start = performance.now();
    if (variant === 'matrix') check(range.getValues());
    else for (const batch of range.batches()) check(batch.values);
    const elapsedMs = performance.now() - start;
    assert.equal(cells, 1_000_000);
    const after = process.memoryUsage();
    const maxRssKiB = process.resourceUsage().maxRSS;
    global.gc();
    console.log(
      JSON.stringify({
        variant,
        elapsedMs,
        cells,
        maxBatch,
        beforeHeap: before.heapUsed,
        observedHeap,
        afterHeap: after.heapUsed,
        retainedHeap: process.memoryUsage().heapUsed,
        maxRssKiB,
      }),
    );
  } finally {
    reader.dispose();
  }
} else {
  const output = mode;
  const before = process.argv[3],
    after = process.argv[4];
  if (!output || !before || !after)
    throw new Error(
      'Usage: node scripts/file-reader-batch-benchmark.mjs NEW_REPORT.json BEFORE_BUNDLE AFTER_BUNDLE',
    );
  const bundles = [
    { name: 'before', path: before },
    { name: 'after', path: after },
  ];
  const report = {
    startedAt: new Date().toISOString(),
    node: process.version,
    host: { platform: platform(), arch: arch(), cpu: cpus()[0].model, memoryBytes: totalmem() },
    source: sourceFingerprint(),
    bundles: bundles.map(({ name, path }) => ({
      name,
      sha256: createHash('sha256').update(readFileSync(path)).digest('hex'),
    })),
    scope:
      'Three alternating fresh Node processes per version and mode. One million blank cells in a 1000x1000 range; timed work includes verifying every value. Heap is sampled between batches and RSS is process high-water, including startup. No browser, formulas, occupied pages, process-isolated baseline subtraction or reference devices.',
    samples: [],
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  try {
    for (let round = 1; round <= 3; round++) {
      const combinations = bundles.flatMap((bundle) =>
        ['matrix', 'batches'].map((variant) => ({ bundle, variant })),
      );
      for (const { bundle, variant } of round % 2 ? combinations : combinations.reverse()) {
        const sample = JSON.parse(
          execFileSync(
            process.execPath,
            ['--expose-gc', script, '--sample', variant, bundle.path],
            {
              encoding: 'utf8',
              timeout: 120000,
            },
          ),
        );
        report.samples.push({ round, bundle: bundle.name, ...sample });
      }
    }
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report.samples));
  } catch (error) {
    report.error = String(error);
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    throw error;
  }
}

import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { arch, cpus, platform, totalmem } from 'node:os';
import { pathToFileURL } from 'node:url';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const [beforePath, afterPath, output] = process.argv.slice(2);
if (!beforePath || !afterPath || !output)
  throw new Error(
    'Usage: node scripts/validation-candidates-benchmark.mjs BEFORE_BUNDLE AFTER_BUNDLE NEW_REPORT',
  );
const variants = [
  { name: 'before', path: beforePath },
  { name: 'after', path: afterPath },
];
for (const variant of variants) {
  variant.hash = createHash('sha256').update(readFileSync(variant.path)).digest('hex');
  variant.collect = (await import(pathToFileURL(variant.path))).validationCandidates;
}
const rule = { kind: 'number', operator: 'gte', first: 1, allowBlank: false, policy: 'reject' };
const rect = (startRow, endRow) => ({ startRow, endRow, startColumn: 0, endColumn: 1 });
const fixtures = [
  {
    name: 'disjoint',
    rules: Array.from({ length: 512 }, (_, i) => ({ id: String(i), range: rect(0, 1), rule })),
    ranges: Array.from({ length: 1800 }, () => rect(1, 2)),
    dimensions: { rows: 2, columns: 1 },
    expected: 0,
  },
  {
    name: 'dense',
    rules: [{ id: 'all', range: rect(0, 65536), rule }],
    ranges: [rect(0, 65536)],
    dimensions: { rows: 65536, columns: 1 },
    expected: 65536,
  },
];
const report = {
  startedAt: new Date().toISOString(),
  runtime: process.version,
  host: { platform: platform(), arch: arch(), cpu: cpus()[0].model, memoryBytes: totalmem() },
  source: sourceFingerprint(),
  bundles: variants.map(({ name, hash }) => ({ name, sha256: hash })),
  scope:
    'Paired Node process, alternating order, two warmups then five measured runs per variant and fixture. Includes input validation, candidate enumeration and sorting; excludes engine, IndexedDB, browser scheduling and peak memory.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (const fixture of fixtures) {
  for (let round = -1; round <= 5; round++) {
    const order = round % 2 ? variants : [...variants].reverse();
    for (const variant of order) {
      const started = performance.now();
      const result = await variant.collect(fixture.rules, fixture.ranges, fixture.dimensions);
      const elapsedMs = performance.now() - started;
      if (
        result.positions.length !== fixture.expected ||
        result.positions.some((p, i) => p.row !== i || p.column !== 0)
      )
        throw new Error(`Incorrect ${fixture.name} result for ${variant.name}`);
      if (round > 0)
        report.samples.push({
          fixture: fixture.name,
          variant: variant.name,
          round,
          elapsedMs,
          count: result.positions.length,
        });
    }
  }
}
writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify(report.samples));

import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/merge-lookup-benchmark.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: ['packages/formats/src/generators/merge-index.ts'],
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
});
const { mergeLookup } = await import(
  `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].text).toString('base64')}`
);
const bounds = { startRow: 0, endRow: 5000, startColumn: 0, endColumn: 20 };
const merges = Array.from({ length: 5000 }, (_, i) => ({
  startRow: i,
  endRow: i + 1,
  startColumn: 0,
  endColumn: 2,
})).reverse();
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  scope:
    'Node lookup-only synthetic benchmark; 100,000 positions, 5,000 non-overlapping one-row merges in reverse order. Includes index construction, verifies each lookup against precomputed expected result. Excludes rendering, workbook loading, browser Workers, process peak memory, other merge geometries and hardware.',
  rounds: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (let round = 1; round <= 6; round++) {
  const samples = [];
  try {
    for (const variant of round % 2 ? ['linear', 'indexed'] : ['indexed', 'linear']) {
      const start = performance.now();
      const lookup =
        variant === 'indexed'
          ? mergeLookup(merges, bounds, 5000)
          : (row, column) =>
              merges.find(
                (m) =>
                  row >= m.startRow &&
                  row < m.endRow &&
                  column >= m.startColumn &&
                  column < m.endColumn,
              );
      const values = new Int32Array(100_000);
      for (let row = 0; row < 5000; row++)
        for (let col = 0; col < 20; col++)
          values[row * 20 + col] = lookup(row, col)?.startRow ?? -1;
      const ms = performance.now() - start;
      for (let row = 0; row < 5000; row++)
        for (let col = 0; col < 20; col++)
          if (values[row * 20 + col] !== (col < 2 ? row : -1))
            throw new Error(`Mismatch ${variant} ${row},${col}`);
      samples.push({ variant, ms, verified: values.length });
    }
    report.rounds.push({ round, samples });
    console.log(JSON.stringify({ round, samples }));
  } catch (error) {
    report.rounds.push({ round, samples, error: String(error) });
    throw error;
  } finally {
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();

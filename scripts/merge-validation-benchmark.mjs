import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/merge-validation-benchmark.mjs NEW_REPORT');
const path = resolve('packages/core/src/model.ts'),
  current = readFileSync(path, 'utf8');
const start = current.indexOf('    for (const range of sh.merges) checkBounds(sh, range);');
const end = current.indexOf('\n  }\n  return s;', start);
if (start < 0 || end < 0 || !current.includes("import { validateRect } from './address.js';"))
  throw new Error('Cannot reconstruct previous validation');
const old =
  current.slice(0, start) +
  `    sh.merges.forEach((r, i) => {
      checkBounds(sh, r);
      assert(!sh.merges.slice(0, i).some(x => intersects(x, r)), 'INVALID_ARGUMENT', 'Overlapping merges');
    });` +
  current.slice(end);
const previous = old.replace(
  "import { validateRect } from './address.js';",
  "import { intersects, validateRect } from './address.js';",
);
const modules = {},
  hashes = {};
for (const [name, contents] of Object.entries({ previous, current })) {
  const result = await build({
    entryPoints: [path],
    bundle: true,
    platform: 'node',
    format: 'esm',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /core\/src\/model\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/core/src'),
          }));
        },
      },
    ],
  });
  hashes[name] = createHash('sha256').update(result.outputFiles[0].text).digest('hex');
  modules[name] = await import(
    `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString('base64')}`
  );
}
const fixture = modules.current.createSnapshot({
  sheets: [{ name: 'Data', rows: 10000, columns: 2 }],
});
fixture.sheets[0].merges = Array.from({ length: 10000 }, (_, row) => ({
  startRow: row,
  endRow: row + 1,
  startColumn: 0,
  endColumn: 2,
})).reverse();
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  hashes,
  scope:
    'In-process Node whole-snapshot validation, 10,000 touching one-row merges in reverse order; reconstructed previous pairwise check vs current sweep. Includes JSON validation and schema parsing. Checks every merge and rejects an overlap outside timing. Excludes persistence, file decoding, browser Worker, peak memory and other merge geometries.',
  rounds: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (let round = 1; round <= 6; round++) {
  const samples = [];
  try {
    for (const variant of round % 2 ? ['previous', 'current'] : ['current', 'previous']) {
      const start = performance.now(),
        result = modules[variant].validateSnapshot(fixture),
        ms = performance.now() - start;
      for (let i = 0; i < 10000; i++)
        if (result.sheets[0].merges[i].startRow !== 9999 - i)
          throw new Error(`Mismatch ${variant} ${i}`);
      const invalid = structuredClone(fixture);
      invalid.sheets[0].merges[9999] = invalid.sheets[0].merges[0];
      let rejected = false;
      try {
        modules[variant].validateSnapshot(invalid);
      } catch (error) {
        rejected = error.message === 'Overlapping merges';
      }
      if (!rejected) throw new Error(`${variant} failed to reject overlap`);
      samples.push({ variant, ms, verified: 10000, rejected });
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

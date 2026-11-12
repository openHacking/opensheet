import { performance } from 'node:perf_hooks';
import { pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';

const paths = process.argv.slice(2);
if (paths.length !== 2 || paths.some((path) => !path.startsWith('/')))
  throw new Error('Pass absolute before and after bundle paths');
const modules = await Promise.all(paths.map((path) => import(pathToFileURL(path).href)));
const hashes = await Promise.all(
  paths.map(async (path) =>
    createHash('sha256')
      .update(await readFile(path))
      .digest('hex'),
  ),
);
const cases = [
  { name: 'literal', target: 'alpha', values: ['alpha', 'bravo'] },
  { name: 'numeric-text', target: 1, values: ['1', '2'] },
  { name: 'wildcard', target: 'a*', values: ['alpha', 'bravo'] },
  { name: 'number', target: 1, values: [1, 2] },
];
const results = [];
for (const fixture of cases) {
  const run = (which) => {
    const predicate = modules[which].criterion(fixture.target);
    let matches = 0;
    const start = performance.now();
    for (let i = 0; i < 10_000; i++) if (predicate(fixture.values[i % 2])) matches++;
    const elapsed = performance.now() - start;
    if (matches !== 5_000) throw new Error(`Wrong result for ${fixture.name}`);
    return elapsed;
  };
  for (let i = 0; i < 4; i++) {
    run(0);
    run(1);
  }
  const pairs = [];
  for (let i = 0; i < 10; i++) {
    const times = [];
    for (const which of i % 2 ? [1, 0] : [0, 1]) times[which] = run(which);
    pairs.push({ beforeMs: times[0], afterMs: times[1] });
  }
  results.push({ name: fixture.name, pairs });
}
console.log(
  JSON.stringify(
    { node: process.version, platform: process.platform, arch: process.arch, hashes, results },
    null,
    2,
  ),
);

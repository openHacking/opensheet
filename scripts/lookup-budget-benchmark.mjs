import { performance } from 'node:perf_hooks';
import { pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';

const paths = process.argv.slice(2);
if (paths.length !== 2 || paths.some((path) => !path.startsWith('/')))
  throw new Error('Pass absolute before and after formula bundle paths');
const modules = await Promise.all(paths.map((path) => import(pathToFileURL(path).href)));
const hashes = await Promise.all(
  paths.map(async (path) =>
    createHash('sha256')
      .update(await readFile(path))
      .digest('hex'),
  ),
);
const keys = Array.from({ length: 5000 }, (_, i) => `key${String(i).padStart(5, '0')}`);
const fixtures = [
  { name: 'exact', formula: 'XLOOKUP("missing",A1:A5000,B1:B5000,-1,0)', expected: -1 },
  { name: 'wildcard', formula: 'MATCH("missing*",A1:A5000,0)', expected: { error: '#N/A' } },
  { name: 'ordered', formula: 'MATCH("key99999",A1:A5000,1)', expected: 5000 },
  { name: 'binary', formula: 'XLOOKUP("key99999",A1:A5000,B1:B5000,-1,0,2)', expected: -1 },
];
const results = [];
for (const fixture of fixtures) {
  const ast = modules.map((m) => m.parseFormula(fixture.formula));
  const read = ({ row, column }) => (column === 0 ? keys[row] : row);
  const run = (which) => {
    const start = performance.now();
    const value = modules[which].evaluateFormula(ast[which], read);
    const elapsed = performance.now() - start;
    if (JSON.stringify(value) !== JSON.stringify(fixture.expected))
      throw new Error(`Wrong result for ${fixture.name}`);
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

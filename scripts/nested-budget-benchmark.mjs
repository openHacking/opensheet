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
const fixtures = [
  {
    name: 'scalar',
    expression: 'A1+B1*2',
    loops: 10_000,
    expected: 5,
    read: ({ column }) => (column === 0 ? 1 : 2),
  },
  {
    name: 'nested-criteria',
    expression: 'SUM(COUNTIF(A1:A1000,"foo"),COUNTIF(A1:A1000,"bar"))',
    loops: 3,
    expected: 1000,
    read: ({ row }) => (row % 2 ? 'foo' : 'bar'),
  },
];
const results = [];
for (const fixture of fixtures) {
  const ast = modules.map((m) => m.parseFormula(fixture.expression));
  const run = (which) => {
    const start = performance.now();
    for (let i = 0; i < fixture.loops; i++) {
      if (modules[which].evaluateFormula(ast[which], fixture.read) !== fixture.expected)
        throw new Error(`Wrong result for ${fixture.name}`);
    }
    return performance.now() - start;
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
  results.push({ name: fixture.name, loops: fixture.loops, pairs });
}
console.log(
  JSON.stringify(
    { node: process.version, platform: process.platform, arch: process.arch, hashes, results },
    null,
    2,
  ),
);

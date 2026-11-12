import { performance } from 'node:perf_hooks';
import { pathToFileURL } from 'node:url';

const path = process.argv[2];
if (!path?.startsWith('/')) throw new Error('Pass an absolute bundled module path');
const { cycleFormulaReference } = await import(pathToFileURL(path).href);
const results = [];
for (const count of [64, 5000, 10000]) {
  const source = `=SUM(${Array(count).fill('A1').join(',')},B2)`;
  const position = source.lastIndexOf('B2') + 1;
  const samplesMs = [];
  for (let i = 0; i < 120; i++) {
    const start = performance.now();
    const edit = cycleFormulaReference(source, position, position);
    const elapsed = performance.now() - start;
    if (edit?.text !== '$B$2') throw new Error('Unexpected reference edit');
    if (i >= 20) samplesMs.push(elapsed);
  }
  results.push({ count, textUnits: source.length, samplesMs });
}
console.log(
  JSON.stringify(
    { node: process.version, platform: process.platform, arch: process.arch, results },
    null,
    2,
  ),
);

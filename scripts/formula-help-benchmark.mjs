import { writeFileSync, readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { cpus, totalmem } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
import { formulaHelp } from '../packages/formula/dist/index.js';
const output = process.argv[2];
if (!output) throw new Error('Pass a new JSON output path');
const fixtures = [
  ['ordinary', '=SUM(1,IF(1,RO'],
  ['deep', '=' + '('.repeat(16000) + 'SUM(' + ')'.repeat(16000)],
  ['long-text', '=CONCAT("' + 'x'.repeat(32750)],
];
const results = fixtures.map(([name, source]) => {
  const cursor = name === 'deep' ? 16005 : source.length;
  const verify = () => {
    const help = formulaHelp(source, cursor);
    if (name === 'ordinary') assert.equal(help.completion.functions[0].name, 'ROUND');
    else if (name === 'deep') assert.equal(help.call.definition.name, 'SUM');
    else assert.equal(help.completion, undefined);
  };
  for (let i = 0; i < 30; i++) verify();
  const samples = [];
  for (let i = 0; i < 200; i++) {
    const start = performance.now();
    formulaHelp(source, cursor);
    samples.push(performance.now() - start);
  }
  verify();
  const sorted = [...samples].sort((a, b) => a - b);
  return {
    name,
    length: source.length,
    cursor,
    samplesMs: samples,
    p50Ms: sorted[99],
    p95Ms: sorted[189],
    p99Ms: sorted[197],
  };
});
writeFileSync(
  output,
  JSON.stringify(
    {
      date: new Date().toISOString(),
      node: process.version,
      cpu: cpus()[0]?.model,
      logicalCPUs: cpus().length,
      memoryBytes: totalmem(),
      buildSha256: createHash('sha256')
        .update(readFileSync('packages/formula/dist/index.js'))
        .digest('hex'),
      scope:
        'Synchronous help extraction only; warmed Node process, excludes browser DOM/paint and Worker calculation. No cross-platform latency guarantee.',
      results,
    },
    null,
    2,
  ) + '\n',
  { flag: 'wx' },
);
console.log(
  JSON.stringify(results.map(({ name, p50Ms, p95Ms, p99Ms }) => ({ name, p50Ms, p95Ms, p99Ms }))),
);

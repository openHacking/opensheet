import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, totalmem } from 'node:os';
import { performance } from 'node:perf_hooks';
import { formulaReferenceHints } from '../packages/formula/dist/index.js';

const output = process.argv[2];
if (!output) throw new Error('Pass a new JSON output path');
const fixtures = [
  { name: 'ordinary', text: '=SUM(B2,D3)', count: 2 },
  { name: 'reference-cap', text: '=' + 'A1+'.repeat(100), count: 64 },
  { name: 'long-prefix', text: '=' + '1+'.repeat(16000) + 'A1', count: 1 },
];
const results = fixtures.map(({ name, text, count }) => {
  const verify = () => assert.equal(formulaReferenceHints(text).length, count);
  for (let i = 0; i < 30; i++) verify();
  const samplesMs = [];
  for (let i = 0; i < 200; i++) {
    const start = performance.now();
    formulaReferenceHints(text);
    samplesMs.push(performance.now() - start);
  }
  verify();
  const sorted = [...samplesMs].sort((a, b) => a - b);
  return {
    name,
    length: text.length,
    samplesMs,
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
        'Warmed synchronous extraction in Node only; excludes DOM, canvas painting, sheet lookup and Worker calculations.',
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

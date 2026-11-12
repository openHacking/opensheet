import { writeFileSync } from 'node:fs';
import { cpus, totalmem } from 'node:os';
import { performance } from 'node:perf_hooks';
import { formatValue } from '../packages/core/dist/index.js';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new output path');
const count = 50000,
  rounds = 5;
const formats = ['0.00', '#,##0.00', '0.0%', '$0.00', 'yyyy-mm-dd hh:mm', 'm/d/yy h:mm'];
const report = {
  date: new Date().toISOString(),
  node: process.version,
  cpu: cpus()[0]?.model,
  logicalCores: cpus().length,
  memoryBytes: totalmem(),
  powerMode: 'unrecorded',
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  count,
  rounds,
  scope:
    'Warmed Node process; 50,000 synchronous formatted values per sample. Does not include Worker transport, browser paint, workbook reads, memory or interactive latency.',
  results: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
try {
  for (const format of formats) {
    const value = (i) =>
      format.includes('yy') ? 45293 + (i % 1440) / 1440 : ((i * 7919) % 1000000) / 100;
    for (let i = 0; i < 5000; i++) formatValue(value(i), format);
    const samples = [];
    let expected;
    for (let round = 0; round < rounds; round++) {
      let digest = 2166136261;
      const start = performance.now();
      for (let i = 0; i < count; i++) {
        const text = formatValue(value(i), format);
        for (let j = 0; j < text.length; j++)
          digest = Math.imul(digest ^ text.charCodeAt(j), 16777619) >>> 0;
      }
      const durationMs = performance.now() - start;
      if (expected !== undefined && digest !== expected)
        throw new Error('Non-deterministic formatting');
      expected = digest;
      samples.push(durationMs);
    }
    report.results.push({ format, checksum: expected, samplesMs: samples });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    process.stdout.write(JSON.stringify(report.results.at(-1)) + '\n');
  }
  report.finishedAt = new Date().toISOString();
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
} catch (error) {
  report.error = String(error);
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  throw error;
}

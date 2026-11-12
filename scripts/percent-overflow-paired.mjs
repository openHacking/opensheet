import { transform } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/percent-overflow-paired.mjs NEW_REPORT');
const source = readFileSync('packages/core/src/format.ts', 'utf8');
const currentBlock = `  const percent = format.includes('%'),
    scaled = percent ? value * 100 : value,
    overflow = percent && !Number.isFinite(scaled);
  let text = numberFormatter(decimals, format.includes(','), overflow).format(
    overflow ? value : scaled,
  );
  if (percent && !overflow) text += '%';`;
const previousBlock = `  let text = numberFormatter(decimals, format.includes(',')).format(
    format.includes('%') ? value * 100 : value,
  );
  if (format.includes('%')) text += '%';`;
if (!source.includes(currentBlock))
  throw new Error('Current formatter has changed; update the control');
const previous = source.replace(currentBlock, previousBlock);
const load = async (text) => {
  const { code } = await transform(text, { loader: 'ts', format: 'esm' });
  return (await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`))
    .formatValue;
};
const versions = { previous: await load(previous), current: await load(source) };
const cases = [
  { name: 'percent', format: '#,##0.00%', value: (i) => ((i * 7919) % 1000000) / 100 },
  { name: 'number', format: '#,##0.00', value: (i) => ((i * 7919) % 1000000) / 100 },
  { name: 'overflow', format: '0.00%', value: (i) => (i % 2 ? -1 : 1) * 1e308 },
];
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  sourceHash: createHash('sha256').update(source).digest('hex'),
  controlHash: createHash('sha256').update(previous).digest('hex'),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  count: 30000,
  scope:
    'Single warmed Node process; control reconstructs only the previous formatValue numeric branch with the current formatter cache. Normal outputs compared for every input. Overflow control is intentionally incorrect. Excludes Worker, paint, process memory, other engines and devices.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (const item of cases) {
  if (item.name !== 'overflow') {
    for (let i = 0; i < report.count; i++)
      if (
        versions.previous(item.value(i), item.format) !==
        versions.current(item.value(i), item.format)
      )
        throw new Error('Normal output changed');
  } else if (
    !versions.previous(item.value(0), item.format).includes('∞') ||
    versions.current(item.value(0), item.format).includes('∞')
  )
    throw new Error('Overflow boundary not exercised');
  for (const version of Object.values(versions))
    for (let i = 0; i < 3000; i++) version(item.value(i), item.format);
  for (let round = 0; round < 6; round++)
    for (const name of round % 2 ? ['current', 'previous'] : ['previous', 'current']) {
      let checksum = 0;
      const start = performance.now();
      for (let i = 0; i < report.count; i++)
        checksum += versions[name](item.value(i), item.format).length;
      report.samples.push({
        case: item.name,
        round: round + 1,
        version: name,
        elapsedMs: performance.now() - start,
        checksum,
      });
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    }
}
report.finishedAt = new Date().toISOString();
writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify(report.samples));

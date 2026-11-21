import { execFileSync } from 'node:child_process';
import { readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/audit-npm-versions.mjs OUTPUT.json');
const registry = 'https://registry.npmjs.org/';
const packages = readdirSync('packages', { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => JSON.parse(readFileSync(join('packages', entry.name, 'package.json'), 'utf8')))
  .map(({ name, version }) => ({ name, candidate: version }))
  .sort((a, b) => a.name.localeCompare(b.name));
if (
  !packages.length ||
  packages.some(
    ({ name, candidate }) => typeof name !== 'string' || typeof candidate !== 'string',
  ) ||
  new Set(packages.map(({ name }) => name)).size !== packages.length
)
  throw new Error('Invalid or empty workspace package inventory');
const records = packages.map(({ name, candidate }) => {
  try {
    const output = execFileSync(
      'npm',
      ['view', name, 'versions', '--json', `--registry=${registry}`],
      {
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'pipe'],
      },
    );
    const result = JSON.parse(output);
    const versions = typeof result === 'string' ? [result] : result;
    if (!Array.isArray(versions) || versions.some((version) => typeof version !== 'string'))
      throw new Error('Unexpected npm versions response');
    return {
      name,
      candidate,
      versions,
      previewPublished: versions.includes('0.1.1'),
      candidatePublished: versions.includes(candidate),
    };
  } catch (error) {
    return { name, candidate, error: error instanceof Error ? error.message : String(error) };
  }
});
const report = {
  date: new Date().toISOString(),
  registry,
  sourceRevision: execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(),
  sourceDirty: !!execFileSync('git', ['status', '--porcelain'], { encoding: 'utf8' }).trim(),
  node: process.version,
  records,
};
writeFileSync(output, `${JSON.stringify(report, null, 2)}\n`);
if (
  records.some(
    (record) => 'error' in record || !record.previewPublished || record.candidatePublished,
  )
)
  process.exitCode = 1;

import { spawnSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';

it('rejects traversal and preserves existing candidate archives', () => {
  const dir = mkdtempSync(join(tmpdir(), 'opensheet-pack-test-'));
  const archive = join(dir, 'existing.tgz');
  writeFileSync(archive, 'keep');
  try {
    const invalid = spawnSync(
      process.execPath,
      ['scripts/pack-candidate.mjs', '..', join(dir, 'new.tgz')],
      { encoding: 'utf8' },
    );
    expect(invalid.status).not.toBe(0);
    const existing = spawnSync(process.execPath, ['scripts/pack-candidate.mjs', 'core', archive], {
      encoding: 'utf8',
    });
    expect(existing.status).not.toBe(0);
    expect(existing.stderr).toContain('Refusing to overwrite archive');
    expect(readFileSync(archive, 'utf8')).toBe('keep');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

it('does not create unverified artifacts or overwrite an existing directory', () => {
  const dir = mkdtempSync(join(tmpdir(), 'opensheet-artifact-test-'));
  const clean = join(dir, 'clean.json'),
    consumer = join(dir, 'consumer.json');
  const version = JSON.parse(readFileSync('package.json', 'utf8')).version;
  const names = readdirSync('packages', { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name);
  const packages = names.map((name) => ({ name, sha256: '0'.repeat(64) }));
  writeFileSync(consumer, JSON.stringify({ version, consumer: { types: 'passed' }, packages }));
  writeFileSync(clean, JSON.stringify({ identical: false, runs: [{ packages }, { packages }] }));
  const output = join(dir, 'new');
  try {
    const invalid = spawnSync(
      process.execPath,
      ['scripts/create-candidate-artifacts.mjs', output, clean, consumer],
      { encoding: 'utf8' },
    );
    expect(invalid.status).not.toBe(0);
    expect(existsSync(output)).toBe(false);
    writeFileSync(clean, JSON.stringify({ identical: true, runs: [{ packages }, { packages }] }));
    mkdirSync(output);
    writeFileSync(join(output, 'keep'), 'sentinel');
    const existing = spawnSync(
      process.execPath,
      ['scripts/create-candidate-artifacts.mjs', output, clean, consumer],
      { encoding: 'utf8' },
    );
    expect(existing.status).not.toBe(0);
    expect(readFileSync(join(output, 'keep'), 'utf8')).toBe('sentinel');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

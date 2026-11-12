import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { report } from '../packages/adapter-sheetjs/src/report.js';

it('keeps workspace packages, generated API version and adapter reports aligned', () => {
  const readJson = (path: string) => JSON.parse(readFileSync(path, 'utf8'));
  const version = readJson('package.json').version;
  for (const entry of readdirSync('packages', { withFileTypes: true })) {
    if (!entry.isDirectory()) continue;
    const manifest = join('packages', entry.name, 'package.json');
    expect(readJson(manifest).version, manifest).toBe(version);
  }
  expect(readJson('api-manifest.json').version).toBe(version);
  expect(report().adapterVersion).toBe(version);
  expect(readFileSync('apps/playground/src/page.ts', 'utf8')).toContain(`v${version}`);
});

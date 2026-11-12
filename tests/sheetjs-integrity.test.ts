import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';

it('pins the direct SheetJS tarball to its reviewed SHA-512 digest', () => {
  const manifest = JSON.parse(readFileSync('package.json', 'utf8'));
  const url = 'https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz';
  const integrity =
    'sha512-oLDq3jw7AcLqKWH2AhCpVTZl8mf6X2YReP+Neh0SJUzV/BdZYjth94tG5toiMB1PPrYtxOCfaoUCkvtuH+3AJA==';
  expect(manifest.devDependencies.xlsx).toBe(url);
  const lock = readFileSync('pnpm-lock.yaml', 'utf8');
  expect(lock).toContain(`resolution: {integrity: ${integrity}, tarball: ${url}}`);
});

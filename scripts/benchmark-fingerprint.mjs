import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

function fingerprint(paths) {
  const files = [...new Set(paths)].sort();
  const hash = createHash('sha256');
  for (const path of files) {
    const content = readFileSync(path);
    hash.update(`${path.length}:${path}:${content.length}:`);
    hash.update(content);
  }
  return {
    algorithm: 'sha256-path-length-content-v1',
    files: files.length,
    sha256: hash.digest('hex'),
  };
}

export function sourceFingerprint() {
  const paths = execFileSync(
    'git',
    ['ls-files', '--cached', '--others', '--exclude-standard', '-z'],
    { encoding: 'utf8' },
  )
    .split('\0')
    .filter(Boolean);
  return fingerprint(
    paths.filter(
      (path) =>
        existsSync(path) &&
        (/^(packages|apps|scripts)\//.test(path) ||
          /^(package\.json|pnpm-lock\.yaml|pnpm-workspace\.yaml|tsconfig\.json)$/.test(path)),
    ),
  );
}

export function localBuildFingerprint(directory = 'apps/playground/dist') {
  if (!existsSync(directory)) return null;
  const files = [];
  const walk = (path) => {
    for (const entry of readdirSync(path, { withFileTypes: true })) {
      const child = join(path, entry.name);
      if (entry.isDirectory()) walk(child);
      else if (entry.isFile()) files.push(child);
    }
  };
  walk(directory);
  return fingerprint(files);
}

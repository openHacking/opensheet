import { execFileSync } from 'node:child_process';
import {
  cpSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, relative, resolve, sep } from 'node:path';

const [folder, output] = process.argv.slice(2);
if (!folder || !output || basename(folder) !== folder || !/^[a-z0-9][a-z0-9-]*$/.test(folder))
  throw new Error('Usage: node scripts/pack-candidate.mjs PACKAGE_DIRECTORY NEW_ARCHIVE.tgz');
const root = resolve('packages');
const packageDir = join(root, folder);
const destination = resolve(output);
if (existsSync(destination)) throw new Error('Refusing to overwrite archive');
const versions = new Map();
for (const entry of readdirSync(root, { withFileTypes: true })) {
  if (!entry.isDirectory()) continue;
  const manifest = JSON.parse(readFileSync(join(root, entry.name, 'package.json'), 'utf8'));
  versions.set(manifest.name, manifest.version);
}
const manifest = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'));
for (const section of [
  'dependencies',
  'devDependencies',
  'optionalDependencies',
  'peerDependencies',
]) {
  for (const [name, spec] of Object.entries(manifest[section] ?? {})) {
    if (typeof spec !== 'string' || !spec.startsWith('workspace:')) continue;
    if (spec !== 'workspace:*' || !versions.has(name))
      throw new Error(`Unsupported workspace dependency ${name}: ${spec}`);
    manifest[section][name] = versions.get(name);
  }
}
const staging = mkdtempSync(join(tmpdir(), 'opensheet-pack-staging-'));
try {
  cpSync(packageDir, staging, {
    recursive: true,
    filter: (path) => {
      const first = relative(packageDir, path).split(sep)[0];
      return first !== 'node_modules' && first !== '.git';
    },
  });
  writeFileSync(join(staging, 'package.json'), JSON.stringify(manifest, null, 2) + '\n');
  execFileSync('pnpm', ['pack', '--out', destination], {
    cwd: staging,
    stdio: 'inherit',
  });
} finally {
  rmSync(staging, { recursive: true, force: true });
}

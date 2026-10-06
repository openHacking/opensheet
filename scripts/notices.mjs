import { createRequire } from 'node:module';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
const appRequire = createRequire(new URL('../package.json', import.meta.url));
const coreRequire = createRequire(new URL('../packages/core/package.json', import.meta.url));
const sections = [];
for (const [name, loader] of [
  ['xlsx', appRequire],
  ['immer', coreRequire],
  ['zod', coreRequire],
]) {
  let dir = dirname(loader.resolve(name));
  while (
    !existsSync(join(dir, 'LICENSE')) &&
    !existsSync(join(dir, 'LICENSE.txt')) &&
    dirname(dir) !== dir
  )
    dir = dirname(dir);
  const path = ['LICENSE', 'LICENSE.txt'].map((n) => join(dir, n)).find(existsSync);
  if (!path) throw new Error(`Missing license for ${name}`);
  sections.push(`${name}\n${'='.repeat(60)}\n${readFileSync(path, 'utf8')}`);
}
mkdirSync('apps/playground/public', { recursive: true });
writeFileSync('apps/playground/public/third-party-licenses.txt', sections.join('\n\n'));

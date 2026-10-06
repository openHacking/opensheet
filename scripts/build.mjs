import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { readdirSync, mkdirSync, copyFileSync, rmSync } from 'node:fs';
for (const name of [
  'formula',
  'core',
  'renderer',
  'plugin-sdk',
  'opensheet',
  'adapter-sheetjs',
  'formats',
  'react',
  'vue',
  'testing',
]) {
  const root = `packages/${name}`;
  rmSync(`${root}/dist`, { recursive: true, force: true });
  mkdirSync(`${root}/dist`, { recursive: true });
  copyFileSync('LICENSE', `${root}/LICENSE`);
  await build({
    entryPoints: [`${root}/src/index.ts`],
    outfile: `${root}/dist/index.js`,
    bundle: true,
    format: 'esm',
    platform: 'neutral',
    target: 'es2022',
    packages: 'external',
    external: ['@opensheetjs/*', 'opensheet', 'react', 'vue', 'immer', 'zod'],
    sourcemap: true,
  });
  execFileSync('pnpm', ['exec', 'tsc', '--project', `${root}/tsconfig.build.json`], {
    stdio: 'inherit',
  });
  if (name === 'opensheet') copyFileSync(`${root}/src/style.css`, `${root}/dist/style.css`);
}

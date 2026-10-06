import { fileURLToPath } from 'node:url';
const names = [
  'core',
  'formula',
  'renderer',
  'plugin-sdk',
  'adapter-sheetjs',
  'formats',
  'react',
  'vue',
  'testing',
];
export const aliases = Object.fromEntries([
  ...names.map((n) => [
    `@opensheetjs/${n}`,
    fileURLToPath(new URL(`../packages/${n}/src/index.ts`, import.meta.url)),
  ]),
  [
    'opensheet/style.css',
    fileURLToPath(new URL('../packages/opensheet/src/style.css', import.meta.url)),
  ],
  ['opensheet', fileURLToPath(new URL('../packages/opensheet/src/index.ts', import.meta.url))],
]);

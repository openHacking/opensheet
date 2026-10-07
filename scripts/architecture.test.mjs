import assert from 'node:assert/strict';
import test from 'node:test';
import { analyzeArchitecture } from './architecture.mjs';
const check = (files) => analyzeArchitecture(new Map(Object.entries(files)));
test('runtime cycles include dynamic imports and re-exports', () => {
  assert.match(
    check({
      'packages/core/src/a.ts': "export * from './b.js';",
      'packages/core/src/b.ts': "const module = import('./a.js');",
    }).join('\n'),
    /Runtime cycle/,
  );
});
test('type-only relationships do not create runtime cycles', () => {
  assert.deepEqual(
    check({
      'packages/core/src/a.ts': "import type { B } from './b.js'; export class A {}",
      'packages/core/src/b.ts': "import { type A } from './a.js'; export type B = A;",
    }),
    [],
  );
});
test('source imports and reverse imports through entry points are rejected', () => {
  const errors = check({
    'packages/core/src/index.ts': "export * from './workbook.js';",
    'packages/core/src/workbook.ts':
      "import { Workbook } from '@opensheetjs/core'; import type { AST } from '../../formula/src/index.js';",
    'packages/formula/src/index.ts': 'export type AST = unknown;',
  }).join('\n');
  assert.match(errors, /internal module imports/);
  assert.match(errors, /cross-package source import/);
});
test('headless layers reject UI packages, frameworks and subpaths', () => {
  const errors = check({
    'packages/core/src/index.ts':
      "import 'react'; import '@opensheetjs/renderer'; import 'opensheet/src/editor';",
    'packages/renderer/src/index.ts': 'export class CanvasGrid {}',
  }).join('\n');
  assert.match(errors, /browser\/framework dependency/);
  assert.match(errors, /forbidden dependency/);
  assert.match(errors, /package subpath/);
});

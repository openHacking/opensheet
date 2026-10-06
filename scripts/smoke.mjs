import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { createWorkbook } from '../packages/core/dist/index.js';
import { createOpenSheet } from '../packages/opensheet/dist/index.js';
import { OpenSheetView as ReactView } from '../packages/react/dist/index.js';
import { OpenSheetView as VueView } from '../packages/vue/dist/index.js';
import { toSheetJS, fromSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { exportRange } from '../packages/formats/dist/index.js';
import { workbookFromRows } from '../packages/testing/dist/index.js';
import { definePlugin } from '../packages/plugin-sdk/dist/index.js';
execFileSync(
  'pnpm',
  [
    'exec',
    'tsc',
    '--noEmit',
    '--module',
    'NodeNext',
    '--moduleResolution',
    'NodeNext',
    '--target',
    'ES2022',
    '--strict',
    'scripts/consumer.mts',
  ],
  { stdio: 'inherit' },
);
assert.equal(typeof globalThis.document, 'undefined');
const book = createWorkbook(),
  sheet = book.getSheets()[0];
sheet.range('A1:B1').setValues([[1, 2]]);
sheet.range('C1').setFormulas([['SUM(A1:B1)']]);
assert.deepEqual(sheet.range('C1').getValues(), [[3]]);
const restored = fromSheetJS(toSheetJS(book.toJSON()).workbook);
assert.equal(restored.snapshot.sheets.length, 1);
assert.match(exportRange(book.toJSON(), { sheetId: sheet.id, format: 'latex' }).text, /tabular/);
assert.equal(typeof createOpenSheet, 'function');
assert.equal(typeof ReactView, 'function');
assert.ok(VueView);
assert.equal(typeof definePlugin, 'function');
assert.equal(workbookFromRows([[1]]).getSheets().length, 1);
book.dispose();
console.log(
  'PASS: all built package entry points import headlessly; formulas, adapters and formats work.',
);

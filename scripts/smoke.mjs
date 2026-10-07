import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { createSnapshot, SnapshotReader, key } from '../packages/core/dist/index.js';
import { createOpenSheet } from '../packages/opensheet/dist/index.js';
import { OpenSheetView as ReactView } from '../packages/react/dist/index.js';
import { OpenSheetView as VueView } from '../packages/vue/dist/index.js';
import { toSheetJS, fromSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { exportRange } from '../packages/formats/dist/index.js';
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
const snapshot = createSnapshot(),
  sheet = snapshot.sheets[0];
for (let col = 0; col < 3; col++)
  sheet.cells[key(sheet.rowOrder[0], sheet.columnOrder[col])] = {
    rowId: sheet.rowOrder[0],
    columnId: sheet.columnOrder[col],
    input:
      col === 2
        ? { type: 'formula', expression: 'SUM(A1:B1)' }
        : { type: 'number', value: col + 1 },
  };
const reader = new SnapshotReader(snapshot);
assert.deepEqual(reader.getSheetById(sheet.id).range('C1').getValues(), [[3]]);
assert.equal(fromSheetJS(toSheetJS(snapshot).workbook).snapshot.sheets.length, 1);
assert.match(exportRange(snapshot, { sheetId: sheet.id, format: 'latex' }).text, /tabular/);
assert.equal(typeof createOpenSheet, 'function');
assert.equal(typeof ReactView, 'function');
assert.ok(VueView);
assert.equal(typeof definePlugin, 'function');
reader.dispose();
console.log(
  'PASS: built package entry points import without a DOM; snapshot formulas, adapters and formats work. Worker-backed workbooks are verified in browser tests.',
);

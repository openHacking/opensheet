import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { createWorkbookFile, FileReader, encodeFileCells } from '../packages/core/dist/index.js';
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
const snapshot = createWorkbookFile(),
  sheet = snapshot.sheets[0];
sheet.blocks = [
  {
    row: 0,
    column: 0,
    data: encodeFileCells(
      Array.from({ length: 3 }, (_, col) => ({
        rowId: 0,
        columnId: col,
        input:
          col === 2
            ? { type: 'formula', expression: 'SUM(A1:B1)' }
            : { type: 'number', value: col + 1 },
      })),
      new Map(),
    ),
  },
];
const reader = new FileReader(snapshot);
assert.deepEqual(reader.getSheetById(sheet.id).range('C1').getValues(), [[3]]);
assert.equal(fromSheetJS(toSheetJS(snapshot).workbook).file.sheets.length, 1);
assert.match(exportRange(snapshot, { sheetId: sheet.id, format: 'latex' }).text, /tabular/);
assert.equal(typeof createOpenSheet, 'function');
assert.equal(typeof ReactView, 'function');
assert.ok(VueView);
assert.equal(typeof definePlugin, 'function');
reader.dispose();
console.log(
  'PASS: built package entry points import without a DOM; JSON formulas, adapters and formats work. Worker-backed workbooks are verified in browser tests.',
);

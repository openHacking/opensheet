import { performance } from 'node:perf_hooks';
import { createWorkbook } from '../packages/core/dist/index.js';
import { fromSheetJS, toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
const measures = {};
const mark = (name, fn) => {
  const start = performance.now();
  const result = fn();
  measures[name] = Number((performance.now() - start).toFixed(2));
  return result;
};
const b = createWorkbook({ sheets: [{ name: 'Benchmark', rows: 10000, columns: 20 }] });
const s = b.getSheets()[0];
mark('write_10000_cells_ms', () =>
  s
    .range('A1:T500')
    .setValues(
      Array.from({ length: 500 }, (_, r) => Array.from({ length: 20 }, (_, c) => r * 20 + c)),
    ),
);
mark('single_edit_ms', () => s.range('A1').setValues([[10]]));
mark('undo_ms', () => b.undo());
const snapshot = mark('serialize_ms', () => b.toJSON());
const output = mark('to_sheetjs_ms', () => toSheetJS(snapshot));
mark('from_sheetjs_ms', () => fromSheetJS(output.workbook));
console.log(
  JSON.stringify(
    {
      node: process.version,
      platform: process.platform,
      arch: process.arch,
      cells: 10000,
      measurements: measures,
      heapMiB: Math.round(process.memoryUsage().heapUsed / 1024 / 1024),
      note: 'Single-run diagnostic, not the design document performance acceptance benchmark.',
    },
    null,
    2,
  ),
);

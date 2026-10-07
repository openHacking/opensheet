import { createWorkbook, type WorkbookFile } from '../packages/core/dist/index.js';
import { createOpenSheet, type OpenSheet } from '../packages/opensheet/dist/index.js';
import { fromSheetJS, toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { exportRange } from '../packages/formats/dist/index.js';
import { definePlugin } from '../packages/plugin-sdk/dist/index.js';
import { OpenSheetView as ReactView } from '../packages/react/dist/index.js';
import { OpenSheetView as VueView } from '../packages/vue/dist/index.js';
const book = await createWorkbook();
const metadata = book.getMetadata();
const axis = book.getAxes(metadata.sheets[0].id, 'row', 0, 1);
const batch = await book.scanCells(metadata.sheets[0].id, { limit: 1, fields: ['input'] });
void [axis, batch];
const snapshot: WorkbookFile = await book.toJSON();
const imported = fromSheetJS(toSheetJS(snapshot).workbook);
exportRange(imported.file, { sheetId: book.getSheets()[0].id, format: 'latex' });
const plugin = definePlugin({
  id: 'consumer.test',
  apiVersion: '^0.1.0',
  version: '0.1.0',
  capabilities: [],
  setup() {},
});
const create: (options: Parameters<typeof createOpenSheet>[0]) => OpenSheet = createOpenSheet;
void [plugin, create, ReactView, VueView];

import { createWorkbook, type WorkbookSnapshot } from '../packages/core/dist/index.js';
import { createOpenSheet, type OpenSheet } from '../packages/opensheet/dist/index.js';
import { fromSheetJS, toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { exportRange } from '../packages/formats/dist/index.js';
import { definePlugin } from '../packages/plugin-sdk/dist/index.js';
import { OpenSheetView as ReactView } from '../packages/react/dist/index.js';
import { OpenSheetView as VueView } from '../packages/vue/dist/index.js';
const book = createWorkbook();
const snapshot: WorkbookSnapshot = book.toJSON();
const imported = fromSheetJS(toSheetJS(snapshot).workbook);
exportRange(imported.snapshot, { sheetId: book.getSheets()[0].id, format: 'latex' });
const plugin = definePlugin({
  id: 'consumer.test',
  apiVersion: '^0.1.0',
  version: '0.1.0',
  capabilities: [],
  setup() {},
});
const create: (options: Parameters<typeof createOpenSheet>[0]) => OpenSheet = createOpenSheet;
void [plugin, create, ReactView, VueView];

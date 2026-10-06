import * as XLSX from 'xlsx';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { validateSnapshot } from '@opensheetjs/core';
self.onmessage = (event: MessageEvent<{ buffer: ArrayBuffer; name: string }>) => {
  try {
    const { buffer, name } = event.data;
    if (buffer.byteLength > 20 * 1024 * 1024) throw new Error('Files are limited to 20 MiB.');
    if (/\.json$/i.test(name)) {
      const snapshot = validateSnapshot(JSON.parse(new TextDecoder().decode(buffer)));
      self.postMessage({ snapshot, report: null });
      return;
    }
    if (/\.xlsx$/i.test(name)) {
      const view = new DataView(buffer);
      let eocd = -1;
      for (let i = buffer.byteLength - 22; i >= Math.max(0, buffer.byteLength - 65557); i--)
        if (view.getUint32(i, true) === 0x06054b50) {
          eocd = i;
          break;
        }
      if (eocd < 0) throw new Error('Invalid XLSX ZIP container');
      const count = view.getUint16(eocd + 10, true);
      let position = view.getUint32(eocd + 16, true),
        expanded = 0;
      if (count > 10000) throw new Error('Too many ZIP entries');
      for (let i = 0; i < count; i++) {
        if (position + 46 > buffer.byteLength || view.getUint32(position, true) !== 0x02014b50)
          throw new Error('Invalid ZIP directory');
        expanded += view.getUint32(position + 24, true);
        if (expanded > 200 * 1024 * 1024) throw new Error('Expanded file exceeds 200 MiB');
        position +=
          46 +
          view.getUint16(position + 28, true) +
          view.getUint16(position + 30, true) +
          view.getUint16(position + 32, true);
      }
    }
    const workbook = XLSX.read(buffer, {
      type: 'array',
      cellNF: true,
      dense: true,
      cellFormula: true,
      cellStyles: true,
      bookVBA: true,
    });
    const imported = fromSheetJS(workbook);
    self.postMessage(imported);
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : String(error) });
  }
};

import * as XLSX from 'xlsx';
import { validateZip } from './zip-guard.js';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { workbookFileSchema } from '@opensheetjs/core';
self.onmessage = async (
  event: MessageEvent<{ buffer: ArrayBuffer; name: string; profile?: boolean }>,
) => {
  try {
    const { buffer, name, profile } = event.data;
    const now = () => (profile === true ? performance.now() : 0);
    const started = now();
    if (buffer.byteLength > 20 * 1024 * 1024) throw new Error('Files are limited to 20 MiB.');
    if (/\.json$/i.test(name)) {
      const file = workbookFileSchema.parse(JSON.parse(new TextDecoder().decode(buffer)));
      self.postMessage({ file, report: null });
      return;
    }
    const storedSizes = await validateZip(buffer, /\.xlsx$/i.test(name));
    const header = new DataView(buffer);
    for (const { offset, size } of storedSizes) header.setUint32(offset, size, true);
    const verified = now();
    const workbook = XLSX.read(buffer, {
      type: 'array',
      cellNF: true,
      dense: true,
      cellFormula: true,
      sheetStubs: true,
      cellStyles: true,
      bookVBA: true,
      bookFiles: true,
    });
    const parsed = now();
    const imported = fromSheetJS(workbook);
    const adapted = now();
    self.postMessage(
      profile === true
        ? {
            ...imported,
            profile: {
              zipMs: verified - started,
              parseMs: parsed - verified,
              adaptMs: adapted - parsed,
              totalMs: adapted - started,
            },
          }
        : imported,
    );
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : String(error) });
  }
};

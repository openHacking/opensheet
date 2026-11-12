import * as XLSX from 'xlsx';
import { writeXlsx } from '@opensheetjs/adapter-sheetjs';
import { OpenSheetError, type WorkbookFile } from '@opensheetjs/core';

self.onmessage = (event: MessageEvent<WorkbookFile>) => {
  try {
    const { bytes, report } = writeXlsx(event.data, XLSX);
    const buffer =
      bytes.buffer instanceof ArrayBuffer &&
      bytes.byteOffset === 0 &&
      bytes.byteLength === bytes.buffer.byteLength
        ? bytes.buffer
        : bytes.slice().buffer;
    self.postMessage({ buffer, report }, { transfer: [buffer] });
  } catch (error) {
    self.postMessage({
      error: error instanceof Error ? error.message : String(error),
      report:
        error instanceof OpenSheetError && error.code === 'UNSUPPORTED_FEATURE'
          ? error.details
          : undefined,
    });
  }
};

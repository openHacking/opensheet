import type { Rect } from '@opensheetjs/core';
export type Format = 'csv' | 'tsv' | 'markdown' | 'html' | 'latex';
export type ExportOptions = {
  sheetId: string;
  range?: string | Rect;
  format: Format;
  options?: { safe?: boolean; bom?: boolean; booktabs?: boolean; header?: boolean };
};
export type ExportResult = { text: string; report: string[]; requiredPackages: string[] };

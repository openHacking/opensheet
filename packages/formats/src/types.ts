import type { Rect } from '@opensheetjs/core';
export type Format = 'csv' | 'tsv' | 'markdown' | 'html' | 'latex';
export type ExportOptions = {
  sheetId: string;
  range?: string | Rect;
  format: Format;
  options?: {
    safe?: boolean;
    formulas?: 'values' | 'preserve';
    bom?: boolean;
    booktabs?: boolean;
    header?: boolean;
  };
};
export type ExportResult = { text: string; report: string[]; requiredPackages: string[] };
/** Published only after a delimited stream completes at one workbook revision. */
export type DelimitedExportReport = {
  format: 'csv' | 'tsv';
  revision: number;
  sanitized: number;
  sample: string[];
  truncated: boolean;
  preservedFormulas?: number;
};

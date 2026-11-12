export type SheetJSWorkbook = {
  SheetNames: string[];
  Sheets: Record<string, any>;
  Workbook?: any;
  vbaraw?: unknown;
  /** Optional raw XLSX parts returned by SheetJS with bookFiles enabled. */
  files?: Record<string, { content?: Uint8Array }>;
};
export type Issue = {
  code: string;
  severity: 'info' | 'warning' | 'error';
  sheetId?: string;
  address?: string;
  feature: string;
  action: 'preserved' | 'approximated' | 'dropped' | 'blocked';
  message: string;
  count?: number;
};
export type CompatibilityReport = {
  adapterVersion: string;
  sourceFormat: string;
  summary: { exact: number; approximated: number; dropped: number; blocked: number };
  issues: Issue[];
  /** Occurrences omitted by the 200-group detail cap; summary always includes them. */
  omittedIssueOccurrences?: number;
};
export type AdapterOptions = { unsupported?: 'report' | 'strict' };

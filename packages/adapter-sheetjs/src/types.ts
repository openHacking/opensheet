export type SheetJSWorkbook = {
  SheetNames: string[];
  Sheets: Record<string, any>;
  Workbook?: any;
  vbaraw?: unknown;
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
};
export type AdapterOptions = { unsupported?: 'report' | 'strict' };

import type { Commit, Selection } from '@opensheetjs/core';
export interface OpenSheetOptions {
  container: string | HTMLElement;
  mode?: 'edit' | 'read';
  toolbar?: boolean;
  onError?: (error: unknown) => void;
}
export type EventMap = {
  'workbook:committed': Commit;
  'selection:changed': Selection;
  'plugin:error': unknown;
  'lifecycle:disposed': undefined;
};

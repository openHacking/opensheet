import {
  type CellValue,
  type Command,
  type Commit,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';
export type Capability =
  | 'selection.read'
  | 'workbook.read'
  | 'workbook.write'
  | 'ui.toolbar'
  | 'events';
export interface PluginContext {
  selection: { get(): Selection | null };
  workbook: {
    getSnapshot(): ReturnType<Workbook['toJSON']>;
    readRange(selection: Selection): { values: CellValue[][] };
    execute(command: Command): void;
  };
  ui: {
    toolbar: { add(action: { id: string; label: string; run(): void }): void };
    notify(message: string): void;
  };
  onCommit(fn: (commit: Commit) => void): void;
  addCleanup(fn: () => void): void;
}
export interface Plugin<T = unknown> {
  id: string;
  version: string;
  apiVersion: string;
  requires?: string[];
  capabilities: Capability[];
  setup(context: PluginContext): { api?: T; dispose?(): void } | void;
}
export function definePlugin<T>(plugin: Plugin<T>): Plugin<T> {
  return plugin;
}
export interface PluginHost {
  getWorkbook(): Workbook;
  getSelection(): Selection | null;
  addToolbar(action: { id: string; label: string; run(): void }): () => void;
  notify(message: string): void;
  onError(error: unknown): void;
}

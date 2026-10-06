import {
  OpenSheetError,
  type Workbook,
  type Selection,
  type Commit,
  type Command,
  type CellValue,
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
export class PluginRegistry {
  private installed = new Map<
    string,
    { plugin: Plugin; api: unknown; cleanups: Array<() => void> }
  >();
  constructor(private host: PluginHost) {}
  has(id: string) {
    return this.installed.has(id);
  }
  get<T = unknown>(id: string) {
    return this.installed.get(id)?.api as T | undefined;
  }
  use(plugins: Plugin[]) {
    const pending = new Map<string, Plugin>();
    for (const p of plugins) {
      if (pending.has(p.id) || this.has(p.id))
        throw new OpenSheetError('PLUGIN_CONFLICT', `Duplicate plugin ${p.id}`);
      if (!/^[a-z][\w.-]+$/.test(p.id) || !['0.1.0', '^0.1.0'].includes(p.apiVersion))
        throw new OpenSheetError('PLUGIN_CONFLICT', `Incompatible plugin ${p.id}`);
      pending.set(p.id, p);
    }
    const order: Plugin[] = [],
      visiting = new Set<string>(),
      visited = new Set<string>();
    const visit = (p: Plugin) => {
      if (visited.has(p.id)) return;
      if (visiting.has(p.id))
        throw new OpenSheetError('PLUGIN_CONFLICT', 'Plugin dependency cycle');
      visiting.add(p.id);
      for (const id of p.requires ?? []) {
        if (this.has(id)) continue;
        const dep = pending.get(id);
        if (!dep) throw new OpenSheetError('PLUGIN_CONFLICT', `Missing dependency ${id}`);
        visit(dep);
      }
      visiting.delete(p.id);
      visited.add(p.id);
      order.push(p);
    };
    plugins.forEach(visit);
    const added: string[] = [];
    try {
      for (const p of order) {
        this.install(p);
        added.push(p.id);
      }
    } catch (e) {
      for (const id of added.reverse()) this.remove(id);
      throw e;
    }
  }
  private install(plugin: Plugin) {
    const cleanups: Array<() => void> = [];
    const require = (cap: Capability) => {
      if (!plugin.capabilities.includes(cap))
        throw new OpenSheetError('PLUGIN_CONFLICT', `Plugin ${plugin.id} did not declare ${cap}`);
    };
    const host = this.host;
    const context: PluginContext = {
      selection: {
        get() {
          require('selection.read');
          return host.getSelection();
        },
      },
      workbook: {
        getSnapshot() {
          require('workbook.read');
          return host.getWorkbook().toJSON();
        },
        readRange(s) {
          require('workbook.read');
          const sheet = host.getWorkbook().getSheetById(s.sheetId);
          if (!sheet) throw new OpenSheetError('INVALID_RANGE', 'Unknown sheet');
          return { values: sheet.range(s).getValues() };
        },
        execute(command) {
          require('workbook.write');
          host.getWorkbook().execute(command);
        },
      },
      ui: {
        toolbar: {
          add(action) {
            require('ui.toolbar');
            cleanups.push(
              host.addToolbar({
                ...action,
                run() {
                  try {
                    action.run();
                  } catch (e) {
                    host.onError(e);
                  }
                },
              }),
            );
          },
        },
        notify(message) {
          require('ui.toolbar');
          host.notify(message);
        },
      },
      onCommit(fn) {
        require('events');
        cleanups.push(host.getWorkbook().onCommit(fn));
      },
      addCleanup(fn) {
        cleanups.push(fn);
      },
    };
    try {
      const result = plugin.setup(context);
      if (result?.dispose) cleanups.push(result.dispose);
      this.installed.set(plugin.id, { plugin, api: result?.api, cleanups });
    } catch (e) {
      for (const fn of cleanups.reverse()) {
        try {
          fn();
        } catch (error) {
          host.onError(error);
        }
      }
      host.onError(e);
      throw e;
    }
  }
  remove(id: string) {
    if ([...this.installed.values()].some((x) => x.plugin.requires?.includes(id)))
      throw new OpenSheetError('PLUGIN_CONFLICT', `Other plugins depend on ${id}`);
    const item = this.installed.get(id);
    if (!item) return;
    this.installed.delete(id);
    for (const fn of item.cleanups.reverse()) {
      try {
        fn();
      } catch (e) {
        this.host.onError(e);
      }
    }
  }
  dispose() {
    for (const id of [...this.installed.keys()].reverse()) this.remove(id);
  }
}

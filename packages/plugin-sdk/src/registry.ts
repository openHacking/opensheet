import { OpenSheetError } from '@opensheetjs/core';
import { createPluginContext } from './context.js';
import { orderPlugins } from './dependencies.js';
import type { Plugin, PluginHost } from './types.js';
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
    const order = orderPlugins(plugins, (id) => this.has(id));
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
    const host = this.host;
    const context = createPluginContext(plugin, host, cleanups);
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

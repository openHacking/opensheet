import { OpenSheetError } from '@opensheetjs/core';
import { createPluginContext } from './context.js';
import { orderPlugins } from './dependencies.js';
import type { Plugin, PluginHost } from './types.js';

function snapshotPlugin(plugin: Plugin): Plugin {
  const { id, version, apiVersion, requires, capabilities, setup, ...extra } = plugin;
  const prototype = Object.getPrototypeOf(plugin);
  return {
    ...extra,
    id,
    version,
    apiVersion,
    requires: requires ? [...requires] : undefined,
    capabilities: [...capabilities],
    // Class methods may depend on private fields that exist only on the instance.
    setup: prototype && prototype !== Object.prototype ? setup.bind(plugin) : setup,
  };
}
export class PluginRegistry {
  private mutating = false;
  private installed = new Map<
    string,
    {
      definition: Plugin;
      requires: string[];
      api: unknown;
      cleanups: Array<() => void>;
      active: { value: boolean };
    }
  >();
  constructor(private host: PluginHost) {}
  private transaction<T>(operation: () => T): T {
    if (this.mutating)
      throw new OpenSheetError(
        'PLUGIN_CONFLICT',
        'Plugin registry mutation is already in progress',
      );
    this.mutating = true;
    try {
      return operation();
    } finally {
      this.mutating = false;
    }
  }
  private report(error: unknown) {
    try {
      this.host.onError(error);
    } catch {
      // An error observer must not interrupt cleanup or mask the original error.
    }
  }
  private cleanup(fn: () => void) {
    try {
      const result = (fn as () => unknown)();
      if (result && typeof (result as PromiseLike<unknown>).then === 'function')
        void Promise.resolve(result).catch((error) => this.report(error));
    } catch (error) {
      this.report(error);
    }
  }
  has(id: string) {
    return this.installed.has(id);
  }
  get<T = unknown>(id: string) {
    return this.installed.get(id)?.api as T | undefined;
  }
  /** Return detached declarations in dependency installation order. */
  definitions(): Plugin[] {
    return [...this.installed.values()].map(({ definition }) => ({
      ...definition,
      requires: definition.requires ? [...definition.requires] : undefined,
      capabilities: [...definition.capabilities],
    }));
  }
  use(plugins: Plugin[]) {
    return this.transaction(() => {
      // Keep lifecycle and permission decisions stable if the caller later edits a definition.
      const definitions = plugins.map(snapshotPlugin);
      const order = orderPlugins(definitions, (id) => this.has(id));
      const added: string[] = [];
      try {
        for (const p of order) {
          const id = p.id;
          this.install(p);
          added.push(id);
        }
      } catch (e) {
        // Newly installed plugins cannot be dependencies of earlier installations.
        // Reverse installation order removes dependents before their dependencies.
        for (const id of added.reverse()) this.uninstall(id);
        throw e;
      }
    });
  }
  private install(plugin: Plugin) {
    const id = plugin.id;
    const requires = [...(plugin.requires ?? [])];
    const definition = {
      ...plugin,
      requires: [...requires],
      capabilities: [...plugin.capabilities],
    };
    const cleanups: Array<() => void> = [];
    const active = { value: true };
    const host = this.host;
    const context = createPluginContext(plugin, host, cleanups, active);
    try {
      const result = plugin.setup(context);
      const pending = result as unknown as PromiseLike<{ dispose?(): void } | void>;
      if (pending && typeof pending.then === 'function') {
        void Promise.resolve(pending).then(
          (late) => this.cleanup(() => late?.dispose?.()),
          (error) => this.report(error),
        );
        throw new OpenSheetError('PLUGIN_CONFLICT', 'Plugin setup must be synchronous');
      }
      const dispose = result?.dispose;
      if (dispose) cleanups.push(() => dispose.call(result));
      this.installed.set(id, { definition, requires, api: result?.api, cleanups, active });
    } catch (e) {
      active.value = false;
      for (const fn of cleanups.reverse()) this.cleanup(fn);
      this.report(e);
      throw e;
    }
  }
  remove(id: string) {
    return this.transaction(() => this.removeInstalled(id));
  }
  private removeInstalled(id: string) {
    if ([...this.installed.values()].some((x) => x.requires.includes(id)))
      throw new OpenSheetError('PLUGIN_CONFLICT', `Other plugins depend on ${id}`);
    this.uninstall(id);
  }
  private uninstall(id: string) {
    const item = this.installed.get(id);
    if (!item) return;
    this.installed.delete(id);
    item.active.value = false;
    for (const fn of item.cleanups.reverse()) this.cleanup(fn);
  }
  dispose() {
    return this.transaction(() => {
      // Registry insertion order is a topological order, including across use() calls.
      for (const id of [...this.installed.keys()].reverse()) this.uninstall(id);
    });
  }
}

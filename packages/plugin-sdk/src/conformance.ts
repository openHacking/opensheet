import { PluginRegistry } from './registry.js';
import type { Workbook } from '@opensheetjs/core';
import type { Plugin, PluginHost } from './types.js';

export interface PluginConformanceResult {
  checks: Array<{ name: string; passed: boolean; detail?: string }>;
  passed: boolean;
}

/** Exercise a plugin bundle against an isolated host. This runs plugin code and cleanup. */
export function checkPluginConformance(
  createPlugins: () => Plugin[],
  host: PluginHost,
): PluginConformanceResult {
  const checks: PluginConformanceResult['checks'] = [];
  const check = (name: string, run: () => void) => {
    try {
      run();
      checks.push({ name, passed: true });
    } catch (error) {
      checks.push({
        name,
        passed: false,
        detail: error instanceof Error ? error.message : String(error),
      });
    }
  };
  const assert = (value: unknown, message: string) => {
    if (!value) throw new Error(message);
  };
  const rejects = (run: () => void, text: string) => {
    try {
      run();
    } catch (error) {
      if (error instanceof Error && error.message.includes(text)) return;
      throw error;
    }
    throw new Error(`Expected rejection containing: ${text}`);
  };
  const base = (id: string, requires: string[] = []): Plugin => ({
    id,
    version: '0.1.0',
    apiVersion: '^0.1.0',
    requires,
    capabilities: [],
    setup() {},
  });
  const probeHost: PluginHost = {
    getWorkbook: () => host.getWorkbook(),
    getSelection: () => host.getSelection(),
    addToolbar: (action) => host.addToolbar(action),
    notify: (message) => host.notify(message),
    onError() {},
  };
  check('registry dependency and duplicate handling', () => {
    const registry = new PluginRegistry(probeHost);
    const dependency = base('conformance.dependency');
    const dependent = base('conformance.dependent', [dependency.id]);
    try {
      registry.use([dependent, dependency]);
      assert(
        registry
          .definitions()
          .map((p) => p.id)
          .join(',') === `${dependency.id},${dependent.id}`,
        'Wrong dependency order',
      );
      rejects(() => registry.remove(dependency.id), 'depend');
      rejects(() => registry.use([dependency]), 'Duplicate');
    } finally {
      registry.dispose();
    }
  });
  check('registry cycle, missing dependency and API version handling', () => {
    const registry = new PluginRegistry(probeHost);
    const a = base('conformance.a', ['conformance.b']);
    const b = base('conformance.b', [a.id]);
    rejects(() => registry.use([a, b]), 'cycle');
    rejects(() => registry.use([a]), 'Missing dependency');
    rejects(
      () => registry.use([{ ...base('conformance.version'), apiVersion: '^999.0.0' }]),
      'Incompatible',
    );
    assert(registry.definitions().length === 0, 'Rejected definitions were installed');
  });
  check('registry rollback and disposal', () => {
    const registry = new PluginRegistry(probeHost);
    const calls: string[] = [];
    const dependency = {
      ...base('conformance.rollback'),
      setup() {
        return {
          dispose() {
            calls.push('disposed');
          },
        };
      },
    };
    const failure = {
      ...base('conformance.failure', [dependency.id]),
      setup() {
        throw new Error('conformance failure');
      },
    };
    rejects(() => registry.use([failure, dependency]), 'conformance failure');
    assert(calls.join(',') === 'disposed', 'Rollback did not dispose the dependency');
    assert(registry.definitions().length === 0, 'Rollback left plugins installed');
    registry.use([dependency]);
    registry.dispose();
    assert(calls.join(',') === 'disposed,disposed', 'Disposal did not run exactly once');
  });
  check('plugin repeated installation and declaration restoration', () => {
    let toolbarCount = 0;
    let listenerCount = 0;
    const errors: unknown[] = [];
    const workbooks = new WeakMap<Workbook, Workbook>();
    const wrapped: PluginHost = {
      getWorkbook() {
        const book = host.getWorkbook();
        let proxy = workbooks.get(book);
        if (!proxy) {
          proxy = new Proxy(book, {
            get(target, property) {
              if (property === 'onCommit') {
                return (listener: Parameters<Workbook['onCommit']>[0]) => {
                  const remove = target.onCommit(listener);
                  listenerCount++;
                  return () => {
                    try {
                      remove();
                    } finally {
                      listenerCount--;
                    }
                  };
                };
              }
              const value = Reflect.get(target, property, target);
              return typeof value === 'function' ? value.bind(target) : value;
            },
          });
          workbooks.set(book, proxy);
        }
        return proxy;
      },
      getSelection: () => host.getSelection(),
      notify: (message) => host.notify(message),
      onError(error) {
        errors.push(error);
      },
      addToolbar(action) {
        const remove = host.addToolbar(action);
        toolbarCount++;
        return () => {
          try {
            remove();
          } finally {
            toolbarCount--;
          }
        };
      },
    };
    const registry = new PluginRegistry(wrapped);
    try {
      const plugins = createPlugins();
      assert(plugins.length > 0, 'Provide at least one plugin');
      registry.use(plugins);
      const definitions = registry.definitions();
      registry.dispose();
      assert(toolbarCount === 0, 'Toolbar actions leaked on disposal');
      assert(listenerCount === 0, 'Commit listeners leaked on disposal');
      registry.use(definitions);
      registry.dispose();
      assert(toolbarCount === 0, 'Toolbar actions leaked after restoration');
      assert(listenerCount === 0, 'Commit listeners leaked after restoration');
      assert(errors.length === 0, `Plugin reported ${errors.length} synchronous error(s)`);
    } finally {
      registry.dispose();
    }
  });
  return { checks, passed: checks.every((item) => item.passed) };
}

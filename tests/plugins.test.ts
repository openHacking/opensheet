import { describe, it, expect, vi } from 'vitest';
import {
  PluginRegistry,
  definePlugin,
  type PluginContext,
  type PluginHost,
  type Plugin,
} from '@opensheetjs/plugin-sdk';
import { createWorkbook } from '@opensheetjs/core';
async function host() {
  const book = await createWorkbook(),
    removed = vi.fn(),
    error = vi.fn();
  const h: PluginHost = {
    getWorkbook: () => book,
    getSelection: () => null,
    addToolbar: () => removed,
    notify: vi.fn(),
    onError: error,
  };
  return { registry: new PluginRegistry(h), removed, error, book };
}
describe('plugins', () => {
  it('retains the disposer and its receiver even if the returned object changes', async () => {
    const { registry, book } = await host();
    const calls: string[] = [];
    const result = {
      marker: 'original receiver',
      dispose() {
        calls.push(this.marker);
      },
    };
    const plugin: Plugin = {
      id: 'example.disposer',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup() {
        return result;
      },
    };
    registry.use([plugin]);
    result.dispose = () => calls.push('replacement');
    registry.remove(plugin.id);
    expect(calls).toEqual(['original receiver']);
    await book.close();
  });

  it('uses the captured disposer during failed installation rollback', async () => {
    const { registry, book } = await host();
    const calls: string[] = [];
    let reads = 0;
    const result = {
      marker: 'original receiver',
      get dispose() {
        reads++;
        return function (this: { marker: string }) {
          calls.push(this.marker);
        };
      },
    };
    const first: Plugin = {
      id: 'example.first',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup() {
        return result;
      },
    };
    const failing: Plugin = {
      id: 'example.failure',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: [first.id],
      capabilities: [],
      setup() {
        Object.defineProperty(result, 'dispose', { value: () => calls.push('replacement') });
        throw new Error('failed setup');
      },
    };
    expect(() => registry.use([failing, first])).toThrow('failed setup');
    expect(reads).toBe(1);
    expect(calls).toEqual(['original receiver']);
    expect(registry.definitions()).toEqual([]);
    await book.close();
  });

  it('rolls back and disposes long chains in reverse order across installations', async () => {
    const { registry, book } = await host();
    const count = 2000;
    const cleanup: number[] = [];
    const missingDependencies: number[] = [];
    const chain = Array.from(
      { length: count },
      (_, i): Plugin => ({
        id: `example.chain-${i}`,
        version: '0.1.0',
        apiVersion: '^0.1.0',
        requires: i ? [`example.chain-${i - 1}`] : [],
        capabilities: [],
        setup() {
          return {
            dispose() {
              if (i > 0 && !registry.has(`example.chain-${i - 1}`)) missingDependencies.push(i);
              cleanup.push(i);
            },
          };
        },
      }),
    );
    registry.use(chain.slice(0, 1));
    const fail: Plugin = {
      id: 'example.failure',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: [`example.chain-${count - 1}`],
      capabilities: [],
      setup() {
        throw new Error('expected failure');
      },
    };
    expect(() => registry.use([fail, ...chain.slice(1)])).toThrow('expected failure');
    expect(registry.has(chain[0].id)).toBe(true);
    expect(registry.definitions().map((p) => p.id)).toEqual([chain[0].id]);
    expect(cleanup).toEqual(Array.from({ length: count - 1 }, (_, i) => count - 1 - i));
    cleanup.length = 0;
    registry.use(chain.slice(1));
    expect(() => registry.remove(chain[0].id)).toThrow('depend');
    registry.dispose();
    expect(cleanup).toEqual(Array.from({ length: count }, (_, i) => count - 1 - i));
    expect(missingDependencies).toEqual([]);
    expect(registry.definitions()).toEqual([]);
    await book.close();
  });
  it('captures each declaration getter once while preserving custom setup state', async () => {
    const { registry, book } = await host();
    const reads: Record<string, number> = {};
    const once = <T>(key: string, value: T) => {
      reads[key] = (reads[key] ?? 0) + 1;
      if (reads[key] > 1) throw new Error(`${key} read twice`);
      return value;
    };
    const plugin = {
      get id() {
        return once('id', 'example.getters');
      },
      get version() {
        return once('version', '0.1.0');
      },
      get apiVersion() {
        return once('apiVersion', '^0.1.0');
      },
      get requires() {
        return once('requires', []);
      },
      get capabilities(): Plugin['capabilities'] {
        return once('capabilities', []);
      },
      get setup(): Plugin['setup'] {
        return once('setup', function (this: { customState: number }) {
          return { api: this.customState };
        });
      },
      customState: 42,
    };
    registry.use([plugin]);
    expect(registry.get('example.getters')).toBe(42);
    expect(reads).toEqual({
      id: 1,
      version: 1,
      apiVersion: 1,
      requires: 1,
      capabilities: 1,
      setup: 1,
    });
    registry.dispose();
    await book.close();
  });
  it('installs and restores class plugins with prototype methods and private state', async () => {
    const { registry, book } = await host();
    const disposed = vi.fn();
    class ClassPlugin implements Plugin<{ count: number }> {
      readonly #secret = 41;
      #count = 0;
      get id() {
        return 'example.class';
      }
      get version() {
        return '0.1.0';
      }
      get apiVersion() {
        return '^0.1.0';
      }
      capabilities: Plugin['capabilities'] = [];
      setup() {
        this.#count++;
        return { api: { count: this.#secret + this.#count }, dispose: disposed };
      }
    }
    const plugin = new ClassPlugin();
    registry.use([plugin]);
    expect(registry.get('example.class')).toEqual({ count: 42 });
    const saved = registry.definitions();
    registry.dispose();
    registry.use(saved);
    expect(registry.get('example.class')).toEqual({ count: 43 });
    registry.dispose();
    expect(disposed).toHaveBeenCalledTimes(2);
    await book.close();
  });
  it('keeps class plugin capabilities fixed when the original instance changes', async () => {
    const { registry, book } = await host();
    class ClassPlugin implements Plugin {
      id = 'example.permissions';
      version = '0.1.0';
      apiVersion = '^0.1.0';
      capabilities: Plugin['capabilities'] = [];
      setup(ctx: PluginContext) {
        return { api: () => ctx.workbook.getMetadata() };
      }
    }
    const plugin = new ClassPlugin();
    registry.use([plugin]);
    plugin.capabilities.push('workbook.read');
    expect(registry.get<() => unknown>(plugin.id)).toThrow();
    const saved = registry.definitions();
    expect(saved[0].capabilities).toEqual([]);
    registry.dispose();
    registry.use(saved);
    expect(registry.get<() => unknown>(plugin.id)).toThrow();
    registry.dispose();
    await book.close();
  });
  it('rejects asynchronous setup, cleans registered resources and disposes a late result', async () => {
    const { registry, removed, book } = await host();
    const lateDispose = vi.fn();
    let finish!: (value: { dispose(): void }) => void;
    const pending = new Promise<{ dispose(): void }>((resolve) => {
      finish = resolve;
    });
    const setup = ((ctx: PluginContext) => {
      ctx.ui.toolbar.add({ id: 'async-setup', label: 'Async', run() {} });
      return pending;
    }) as unknown as Plugin['setup'];
    expect(() =>
      registry.use([
        definePlugin({
          id: 'example.async',
          version: '0.1.0',
          apiVersion: '^0.1.0',
          capabilities: ['ui.toolbar'],
          setup,
        }),
      ]),
    ).toThrow('setup must be synchronous');
    expect(removed).toHaveBeenCalledTimes(1);
    expect(registry.has('example.async')).toBe(false);
    finish({ dispose: lateDispose });
    await Promise.resolve();
    await Promise.resolve();
    expect(lateDispose).toHaveBeenCalledTimes(1);
    await book.close();
  });

  it('observes late asynchronous setup and cleanup failures without blocking later cleanup', async () => {
    const { registry, error, book } = await host();
    let rejectSetup!: (error: Error) => void;
    const pending = new Promise<never>((_, reject) => {
      rejectSetup = reject;
    });
    const setup = (() => pending) as unknown as Plugin['setup'];
    expect(() =>
      registry.use([
        definePlugin({
          id: 'example.async-error',
          version: '0.1.0',
          apiVersion: '^0.1.0',
          capabilities: [],
          setup,
        }),
      ]),
    ).toThrow('setup must be synchronous');
    rejectSetup(new Error('late setup failure'));
    await Promise.resolve();
    await Promise.resolve();
    expect(error).toHaveBeenCalledWith(expect.objectContaining({ message: 'late setup failure' }));
    const cleaned = vi.fn();
    registry.use([
      definePlugin({
        id: 'example.async-cleanup',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: [],
        setup(ctx) {
          ctx.addCleanup(cleaned);
          ctx.addCleanup(async () => {
            throw new Error('late cleanup failure');
          });
        },
      }),
    ]);
    registry.dispose();
    expect(cleaned).toHaveBeenCalledTimes(1);
    await Promise.resolve();
    await Promise.resolve();
    expect(error).toHaveBeenCalledWith(
      expect.objectContaining({ message: 'late cleanup failure' }),
    );
    await book.close();
  });
  it('pins identity, dependencies and capabilities before calling setup', async () => {
    const { registry, book } = await host();
    const source = definePlugin({
      id: 'example.anchor',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup() {},
    });
    const dependent = definePlugin({
      id: 'example.mutating',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: ['example.anchor'],
      capabilities: [],
      setup(ctx) {
        this.id = 'example.changed';
        this.requires!.length = 0;
        this.capabilities.push('workbook.read');
        expect(() => ctx.workbook.getMetadata()).toThrow('did not declare');
      },
    });
    registry.use([source, dependent]);
    expect(registry.has('example.mutating')).toBe(true);
    expect(registry.has('example.changed')).toBe(false);
    const snapshot = registry.definitions();
    expect(snapshot.map((plugin) => plugin.id)).toEqual(['example.anchor', 'example.mutating']);
    expect(snapshot[1].requires).toEqual(['example.anchor']);
    expect(snapshot[1].capabilities).toEqual([]);
    snapshot[1].requires!.length = 0;
    expect(() => registry.remove('example.anchor')).toThrow('depend');
    registry.dispose();
    await book.close();
  });

  it('rolls back the original identity when setup mutates its definition', async () => {
    const { registry, book } = await host();
    const cleanup = vi.fn();
    const changing = definePlugin({
      id: 'example.change-id',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup(ctx) {
        this.id = 'example.new-id';
        ctx.addCleanup(cleanup);
      },
    });
    const failing = definePlugin({
      id: 'example.failure',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup() {
        throw new Error('failed');
      },
    });
    expect(() => registry.use([changing, failing])).toThrow('failed');
    expect(registry.has('example.change-id')).toBe(false);
    expect(registry.has('example.new-id')).toBe(false);
    expect(cleanup).toHaveBeenCalledTimes(1);
    await book.close();
  });
  it('keeps dependency and capability declarations stable after caller mutation', async () => {
    const { registry, book } = await host();
    const first = definePlugin({
      id: 'example.source',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup() {},
    });
    let context!: PluginContext;
    const second = definePlugin({
      id: 'example.dependent',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: ['example.source'],
      capabilities: [],
      setup(ctx) {
        context = ctx;
      },
    });
    registry.use([first, second]);
    second.requires!.length = 0;
    second.capabilities.push('workbook.read' as never);
    expect(() => registry.remove('example.source')).toThrow('depend');
    expect(() => context.workbook.getMetadata()).toThrow('declare');
    registry.dispose();
    await book.close();
  });
  it('rejects registry mutations during setup and rolls back earlier dependencies', async () => {
    const { registry, removed, book } = await host();
    const first = definePlugin({
      id: 'example.first',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: ['ui.toolbar'],
      setup(ctx) {
        ctx.ui.toolbar.add({ id: 'first', label: 'First', run() {} });
      },
    });
    const second = definePlugin({
      id: 'example.second',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: ['example.first'],
      capabilities: [],
      setup() {
        registry.remove('example.first');
      },
    });
    expect(() => registry.use([second, first])).toThrow('mutation is already in progress');
    expect(registry.has('example.first')).toBe(false);
    expect(registry.has('example.second')).toBe(false);
    expect(removed).toHaveBeenCalledTimes(1);
    registry.use([first]);
    expect(registry.has('example.first')).toBe(true);
    registry.dispose();
    await book.close();
  });

  it('rejects reentrant installation during cleanup while continuing disposal', async () => {
    const { registry, error, book } = await host();
    const cleaned = vi.fn();
    const late = definePlugin({
      id: 'example.late',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup() {},
    });
    registry.use([
      definePlugin({
        id: 'example.cleanup',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: [],
        setup(ctx) {
          ctx.addCleanup(cleaned);
          ctx.addCleanup(() => registry.use([late]));
        },
      }),
    ]);
    registry.dispose();
    expect(cleaned).toHaveBeenCalledTimes(1);
    expect(error).toHaveBeenCalledWith(expect.objectContaining({ code: 'PLUGIN_CONFLICT' }));
    expect(registry.has('example.late')).toBe(false);
    registry.use([late]);
    expect(registry.has('example.late')).toBe(true);
    registry.dispose();
    await book.close();
  });
  it('invalidates retained contexts and actions after removal and calls disposal with its receiver', async () => {
    const book = await createWorkbook();
    let action!: { run(): void | Promise<void> };
    const removed = vi.fn(),
      run = vi.fn();
    const registry = new PluginRegistry({
      getWorkbook: () => book,
      getSelection: () => null,
      addToolbar: (value) => {
        action = value;
        return removed;
      },
      notify: vi.fn(),
      onError: vi.fn(),
    });
    let context!: PluginContext;
    const instance = {
      disposed: false,
      dispose() {
        this.disposed = true;
      },
    };
    registry.use([
      definePlugin({
        id: 'example.lifetime',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: ['ui.toolbar', 'workbook.read'],
        setup(ctx) {
          context = ctx;
          ctx.ui.toolbar.add({ id: 'action', label: 'Action', run });
          return instance;
        },
      }),
    ]);
    action.run();
    expect(run).toHaveBeenCalledTimes(1);
    registry.remove('example.lifetime');
    expect(instance.disposed).toBe(true);
    expect(removed).toHaveBeenCalledTimes(1);
    expect(() => context.workbook.toJSON()).toThrow('no longer active');
    expect(() => context.ui.toolbar.add({ id: 'late', label: 'Late', run })).toThrow(
      'no longer active',
    );
    expect(() => context.addCleanup(() => {})).toThrow('no longer active');
    action.run();
    expect(run).toHaveBeenCalledTimes(1);
    await book.close();
  });

  it('invalidates a context retained by a failed setup', async () => {
    const { registry, book } = await host();
    let context!: PluginContext;
    expect(() =>
      registry.use([
        definePlugin({
          id: 'example.failed',
          version: '0.1.0',
          apiVersion: '^0.1.0',
          capabilities: ['workbook.read'],
          setup(ctx) {
            context = ctx;
            throw new Error('setup failed');
          },
        }),
      ]),
    ).toThrow('setup failed');
    expect(() => context.workbook.toJSON()).toThrow('no longer active');
    await book.close();
  });

  it('runs all rollback cleanups and preserves the setup error when the error observer throws', async () => {
    const book = await createWorkbook();
    const cleaned = vi.fn();
    const registry = new PluginRegistry({
      getWorkbook: () => book,
      getSelection: () => null,
      addToolbar: () => () => {},
      notify: () => {},
      onError: () => {
        throw new Error('observer failed');
      },
    });
    expect(() =>
      registry.use([
        definePlugin({
          id: 'example.rollback',
          version: '0.1.0',
          apiVersion: '^0.1.0',
          capabilities: [],
          setup(ctx) {
            ctx.addCleanup(cleaned);
            ctx.addCleanup(() => {
              throw new Error('cleanup failed');
            });
            throw new Error('setup failed');
          },
        }),
      ]),
    ).toThrow('setup failed');
    expect(cleaned).toHaveBeenCalledTimes(1);
    expect(registry.has('example.rollback')).toBe(false);
    await book.close();
  });
  it('installs dependencies in order and cleans resources in reverse', async () => {
    const { registry, removed } = await host(),
      order: string[] = [];
    const a = definePlugin({
      id: 'example.a',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: ['ui.toolbar'],
      setup(ctx) {
        order.push('a');
        ctx.ui.toolbar.add({ id: 'a', label: 'A', run() {} });
        return {
          api: { answer: 42 },
          dispose() {
            order.push('dispose-a');
          },
        };
      },
    });
    const b = definePlugin({
      id: 'example.b',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: ['example.a'],
      capabilities: [],
      setup() {
        order.push('b');
        return {
          dispose() {
            order.push('dispose-b');
          },
        };
      },
    });
    registry.use([b, a]);
    expect(order).toEqual(['a', 'b']);
    expect(registry.get('example.a')).toEqual({ answer: 42 });
    expect(() => registry.remove('example.a')).toThrow();
    registry.dispose();
    expect(order.slice(-2)).toEqual(['dispose-b', 'dispose-a']);
    expect(removed).toHaveBeenCalledTimes(1);
  });
  it('rolls back partial installation and setup resources', async () => {
    const { registry, removed } = await host();
    expect(() =>
      registry.use([
        definePlugin({
          id: 'example.bad',
          version: '0.1.0',
          apiVersion: '^0.1.0',
          capabilities: ['ui.toolbar'],
          setup(ctx) {
            ctx.ui.toolbar.add({ id: 'x', label: 'X', run() {} });
            throw new Error('failed');
          },
        }),
      ]),
    ).toThrow();
    expect(registry.has('example.bad')).toBe(false);
    expect(removed).toHaveBeenCalledTimes(1);
  });
  it('rejects missing permissions and incompatible versions', async () => {
    const { registry } = await host();
    expect(() =>
      registry.use([
        definePlugin({
          id: 'example.bad',
          version: '0.1.0',
          apiVersion: '^0.1.0',
          capabilities: [],
          setup(ctx) {
            ctx.workbook.toJSON();
          },
        }),
      ]),
    ).toThrow(/declare/);
    expect(() =>
      registry.use([
        definePlugin({
          id: 'example.new',
          version: '1',
          apiVersion: '^2.0.0',
          capabilities: [],
          setup() {},
        }),
      ]),
    ).toThrow(/Incompatible/);
  });
});

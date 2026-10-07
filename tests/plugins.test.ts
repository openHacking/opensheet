import { describe, it, expect, vi } from 'vitest';
import { PluginRegistry, definePlugin, type PluginHost } from '@opensheetjs/plugin-sdk';
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
            ctx.workbook.getSnapshot();
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

import { describe, expect, it } from 'vitest';
import type { Plugin } from '@opensheetjs/plugin-sdk';
import { orderPlugins } from '../packages/plugin-sdk/src/dependencies.js';

const plugin = (index: number, requires: string[] = []): Plugin => ({
  id: `example.plugin-${index}`,
  version: '0.1.0',
  apiVersion: '^0.1.0',
  requires,
  capabilities: [],
  setup() {},
});

describe('plugin dependency ordering', () => {
  it('orders a deep dependency chain without overflowing the JavaScript call stack', () => {
    const count = 15_000;
    const plugins = Array.from({ length: count }, (_, i) =>
      plugin(i, i + 1 < count ? [`example.plugin-${i + 1}`] : []),
    );
    const ordered = orderPlugins(plugins, () => false);
    expect(ordered.map((p) => p.id)).toEqual(plugins.map((p) => p.id).reverse());
  });

  it('detects a deep cycle and reports the declared conflict', () => {
    const count = 15_000;
    const plugins = Array.from({ length: count }, (_, i) =>
      plugin(i, [`example.plugin-${(i + 1) % count}`]),
    );
    expect(() => orderPlugins(plugins, () => false)).toThrow('Plugin dependency cycle');
  });

  it('preserves input and dependency traversal order with shared and installed dependencies', () => {
    const shared = plugin(0);
    const left = plugin(1, [shared.id, 'example.installed']);
    const right = plugin(2, [shared.id]);
    const root = plugin(3, [right.id, left.id, shared.id]);
    const standalone = plugin(4);
    expect(
      orderPlugins([root, standalone, left, right, shared], (id) => id === 'example.installed'),
    ).toEqual([shared, right, left, root, standalone]);
    expect(() => orderPlugins([plugin(5, ['example.missing'])], () => false)).toThrow(
      'Missing dependency example.missing',
    );
    expect(() => orderPlugins([shared, shared], () => false)).toThrow('Duplicate plugin');
  });
});

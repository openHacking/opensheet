import { describe, expect, it, vi } from 'vitest';
import { checkPluginConformance, type Plugin, type PluginHost } from '@opensheetjs/plugin-sdk';
import { createWorkbook } from '@opensheetjs/core';
import { selectionSum } from '../examples/plugin.js';
import { activityCounter } from '../examples/plugin-activity.js';

describe('plugin conformance helper', () => {
  it('checks both public-entry-point example plugins on a disposable workbook', async () => {
    const workbook = await createWorkbook();
    const actions = new Set<string>();
    const host: PluginHost = {
      getWorkbook: () => workbook,
      getSelection: () => null,
      addToolbar(action) {
        if (actions.has(action.id)) throw new Error('duplicate action');
        actions.add(action.id);
        return () => {
          actions.delete(action.id);
        };
      },
      notify() {},
      onError() {},
    };
    const report = checkPluginConformance(() => [selectionSum, activityCounter], host);
    expect(report).toMatchObject({ passed: true });
    expect(actions.size).toBe(0);
    await workbook.close();
  });
  it('checks a real plugin bundle, including repeated toolbar cleanup', async () => {
    const workbook = await createWorkbook();
    const actions = new Set<string>();
    const host: PluginHost = {
      getWorkbook: () => workbook,
      getSelection: () => null,
      addToolbar(action) {
        if (actions.has(action.id)) throw new Error('duplicate action');
        actions.add(action.id);
        return () => {
          actions.delete(action.id);
        };
      },
      notify() {},
      onError: vi.fn(),
    };
    let disposals = 0;
    const create = (): Plugin[] => [
      {
        id: 'example.action',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: ['ui.toolbar'],
        setup(ctx) {
          ctx.ui.toolbar.add({ id: 'test', label: 'Test', run() {} });
          return {
            dispose() {
              disposals++;
            },
          };
        },
      },
    ];
    const report = checkPluginConformance(create, host);
    expect(report.checks).toHaveLength(4);
    expect(report.passed).toBe(true);
    expect(actions.size).toBe(0);
    expect(disposals).toBe(2);
    expect(host.onError).not.toHaveBeenCalled();
    await workbook.close();
  });

  it('reports plugin setup and synchronous cleanup failures', async () => {
    const workbook = await createWorkbook();
    const host: PluginHost = {
      getWorkbook: () => workbook,
      getSelection: () => null,
      addToolbar: () => () => {},
      notify() {},
      onError() {},
    };
    const plugin: Plugin = {
      id: 'example.bad',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup(ctx) {
        ctx.addCleanup(() => {
          throw new Error('cleanup failed');
        });
      },
    };
    const failed = checkPluginConformance(() => [plugin], host);
    expect(failed.passed).toBe(false);
    expect(failed.checks.at(-1)).toMatchObject({
      passed: false,
      detail: expect.stringContaining('synchronous error'),
    });
    const missing = checkPluginConformance(
      () => [{ ...plugin, requires: ['example.missing'] }],
      host,
    );
    expect(missing.passed).toBe(false);
    expect(missing.checks.at(-1)).toMatchObject({
      passed: false,
      detail: expect.stringContaining('Missing dependency'),
    });
    await workbook.close();
  });

  it('tracks commit-listener cleanup and reports an unsubscribe failure', async () => {
    const workbook = await createWorkbook();
    const subscribe = workbook.onCommit;
    let active = 0;
    let failCleanup = false;
    workbook.onCommit = (callback) => {
      const remove = subscribe.call(workbook, callback);
      active++;
      return () => {
        remove();
        active--;
        if (failCleanup) throw new Error('unsubscribe failed');
      };
    };
    const host: PluginHost = {
      getWorkbook: () => workbook,
      getSelection: () => null,
      addToolbar: () => () => {},
      notify() {},
      onError() {},
    };
    expect(checkPluginConformance(() => [activityCounter], host).passed).toBe(true);
    expect(active).toBe(0);
    failCleanup = true;
    const report = checkPluginConformance(() => [activityCounter], host);
    expect(report.passed).toBe(false);
    expect(report.checks.at(-1)).toMatchObject({
      passed: false,
      detail: expect.stringContaining('synchronous error'),
    });
    expect(active).toBe(0);
    workbook.onCommit = subscribe;
    await workbook.close();
  });
});

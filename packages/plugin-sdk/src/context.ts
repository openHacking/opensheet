import { OpenSheetError } from '@opensheetjs/core';
import type { Capability, Plugin, PluginContext, PluginHost } from './types.js';
export function createPluginContext(
  plugin: Plugin,
  host: PluginHost,
  cleanups: Array<() => void>,
  active: { value: boolean },
): PluginContext {
  const id = plugin.id;
  const capabilities = new Set(plugin.capabilities);
  const assertActive = () => {
    if (!active.value)
      throw new OpenSheetError('PLUGIN_CONFLICT', `Plugin ${id} is no longer active`);
  };
  const report = (error: unknown) => {
    try {
      host.onError(error);
    } catch {
      // Error observers must not turn rejected actions into unhandled rejections.
    }
  };
  const require = (cap: Capability) => {
    assertActive();
    if (!capabilities.has(cap))
      throw new OpenSheetError('PLUGIN_CONFLICT', `Plugin ${id} did not declare ${cap}`);
  };
  const context: PluginContext = {
    selection: {
      get() {
        require('selection.read');
        return host.getSelection();
      },
    },
    workbook: {
      getMetadata() {
        require('workbook.read');
        return host.getWorkbook().getMetadata();
      },
      toJSON() {
        require('workbook.read');
        return host.getWorkbook().toJSON();
      },
      async readRange(s) {
        require('workbook.read');
        const sheet = host.getWorkbook().getSheetById(s.sheetId);
        if (!sheet) throw new OpenSheetError('INVALID_RANGE', 'Unknown sheet');
        return { values: await sheet.range(s).getValues() };
      },
      execute(command) {
        require('workbook.write');
        return host.getWorkbook().execute(command);
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
                if (!active.value) return;
                try {
                  Promise.resolve(action.run()).catch(report);
                } catch (e) {
                  report(e);
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
      assertActive();
      cleanups.push(fn);
    },
  };
  return context;
}

import { OpenSheetError } from '@opensheetjs/core';
import type { Capability, Plugin, PluginContext, PluginHost } from './types.js';
export function createPluginContext(
  plugin: Plugin,
  host: PluginHost,
  cleanups: Array<() => void>,
): PluginContext {
  const require = (cap: Capability) => {
    if (!plugin.capabilities.includes(cap))
      throw new OpenSheetError('PLUGIN_CONFLICT', `Plugin ${plugin.id} did not declare ${cap}`);
  };
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
  return context;
}

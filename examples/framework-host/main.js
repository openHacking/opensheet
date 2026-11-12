import React, { StrictMode, createElement } from 'react';
import { createRoot, hydrateRoot } from 'react-dom/client';
import { createApp, createSSRApp, h, shallowReactive, nextTick } from 'vue';
import { OpenSheetView as ReactView } from '@opensheetjs/react';
import { OpenSheetView as VueView } from '@opensheetjs/vue';
import { OpenSheet, createWorkbookFile, encodeFileCells, openWorkbook } from 'opensheet';
import 'opensheet/style.css';

const framework = new URLSearchParams(location.search).get('framework') ?? 'react';
const state = { ready: 0, changes: [], errors: [], app: null, pending: 0, workerUrls: [] };
state.hydrationErrors = [];
state.hydratedHostPreserved = null;
const workers = new Set();
let createdWorkers = 0;
const NativeWorker = Worker;
window.Worker = class extends NativeWorker {
  constructor(...args) {
    super(...args);
    workers.add(this);
    createdWorkers++;
    state.workerUrls.push(String(args[0]));
  }
  terminate() {
    workers.delete(this);
    super.terminate();
  }
};
let root, props, gate, release;
const frames = () =>
  new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
for (const method of ['createWorkbook', 'load']) {
  const original = OpenSheet.prototype[method];
  OpenSheet.prototype[method] = async function (...args) {
    const wait = gate;
    const result = await original.apply(this, args);
    if (wait) {
      state.pending++;
      await wait;
      state.pending--;
    }
    return result;
  };
}
const file = (value) => {
  const result = createWorkbookFile({ sheets: [{ name: 'Host', rows: 100, columns: 8 }] });
  result.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        [{ rowId: 0, columnId: 0, input: { type: 'number', value } }],
        new Map(),
      ),
    },
  ];
  return result;
};
const callbacks = (version = 0) => ({
  onReady: (app) => {
    state.ready++;
    state.app = app;
  },
  onChange: (commit) => state.changes.push({ version, revision: commit.revision }),
  onError: (error) => state.errors.push({ version, message: String(error) }),
});
const renderReact = () =>
  root.render(createElement(StrictMode, null, createElement(ReactView, props)));
window.hostFixture = {
  state,
  framework,
  mode: import.meta.env.MODE,
  resources: () => ({
    workers: workers.size,
    createdWorkers,
    grids: document.querySelectorAll('.os-grid').length,
  }),
  async mount(options = {}) {
    const container = document.querySelector('#root');
    const hydrate = container.hasAttribute('data-ssr');
    const serverHost = hydrate ? container.firstElementChild : null;
    state.ready = 0;
    state.changes = [];
    state.errors = [];
    state.app = null;
    props = {
      ...callbacks(options.version),
      initialFile: options.invalid ? { schemaVersion: 999 } : file(options.value ?? 1),
      readOnly: !!options.readOnly,
    };
    if (framework === 'react') {
      if (hydrate) {
        root = hydrateRoot(
          container,
          createElement(StrictMode, null, createElement(ReactView, props)),
          {
            onRecoverableError: (error) => state.hydrationErrors.push(String(error)),
          },
        );
      } else {
        root = createRoot(container);
        renderReact();
      }
    } else {
      props = shallowReactive(props);
      root = (hydrate ? createSSRApp : createApp)({ setup: () => () => h(VueView, props) });
      root.mount('#root');
    }
    await nextTick();
    await frames();
    if (hydrate) {
      state.hydratedHostPreserved = serverHost === container.firstElementChild;
      container.removeAttribute('data-ssr');
    }
  },
  async update(options) {
    const next = {
      ...(options.value !== undefined ? { initialFile: file(options.value) } : {}),
      ...(options.readOnly !== undefined ? { readOnly: options.readOnly } : {}),
      ...(options.version !== undefined ? callbacks(options.version) : {}),
    };
    if (framework === 'react') {
      props = { ...props, ...next };
      renderReact();
    } else Object.assign(props, next);
    await nextTick();
    await frames();
  },
  async replace(value) {
    await state.app.load(file(value));
  },
  async reopen(id) {
    const book = await openWorkbook(id);
    try {
      return {
        readOnly: book.isReadOnly,
        value: await book.getSheets()[0].range('A1').getValues(),
      };
    } finally {
      await book.close();
    }
  },
  async unmount() {
    root?.unmount();
    root = undefined;
    state.app = null;
    await frames();
  },
  hold() {
    gate = new Promise((resolve) => {
      release = resolve;
    });
  },
  async release() {
    release?.();
    gate = undefined;
    await frames();
  },
  installLateAction() {
    state.app.use({
      id: 'host.late',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: ['ui.toolbar'],
      setup(ctx) {
        ctx.ui.toolbar.add({
          id: 'host.late.action',
          label: 'Late action',
          run: () =>
            new Promise((_, reject) => {
              window.hostFixture.rejectLate = () => reject(new Error('late host error'));
            }),
        });
      },
    });
  },
};

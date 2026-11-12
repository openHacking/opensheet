import assert from 'node:assert/strict';
import { createElement } from 'react';
import { renderToString as renderReact } from 'react-dom/server';
import { createSSRApp, h } from 'vue';
import { renderToString as renderVue } from 'vue/server-renderer';
import { OpenSheetView as ReactView } from '@opensheetjs/react';
import { OpenSheetView as VueView } from '@opensheetjs/vue';
import { createWorkbookFile, FileReader } from 'opensheet';

for (const name of ['document', 'window', 'Worker', 'indexedDB'])
  assert.equal(typeof globalThis[name], 'undefined', `${name} must not be installed for SSR`);
const initialFile = createWorkbookFile();
const reader = new FileReader(initialFile);
assert.deepEqual(reader.getSheetById(initialFile.sheets[0].id).range('A1').getValues(), [[null]]);
reader.dispose();
let ready = 0;
const props = { initialFile, onReady: () => ready++ };
const react = renderReact(createElement(ReactView, props));
const vue = await renderVue(createSSRApp({ render: () => h(VueView, props) }));
for (const markup of [react, vue]) {
  assert.match(markup, /^<div\b/);
  assert.ok(!markup.includes('os-grid'));
  assert.match(markup, /<\/div>$/);
}
assert.equal(ready, 0, 'Server rendering must not mount an editor');
console.log(JSON.stringify({ react, vue, ready, status: 'passed' }));

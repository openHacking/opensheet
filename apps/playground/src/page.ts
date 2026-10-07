import {
  Braces,
  ChevronDown,
  Code2,
  Copy,
  CornerDownLeft,
  createElement,
  Download,
  FileSpreadsheet,
  HelpCircle,
  RotateCcw,
  Search,
  ShieldCheck,
  Table2,
  Upload,
  X,
  type IconNode,
} from 'lucide';
import { apiExample, pluginExample } from './developer-pages.js';
import { scenes } from './scenes.js';
export function mountPage() {
  const icon = (node: IconNode) =>
    createElement(node, { width: 16, height: 16, 'aria-hidden': 'true', focusable: 'false' })
      .outerHTML;
  const asset = (name: string) => `${import.meta.env.BASE_URL}assets/${name}`;
  const logo = `<img src="${asset('icon.svg')}" width="30" height="30" alt="" />`;
  const sceneOptions = Object.entries(scenes)
    .map(([id, scene]) => `<option value="${id}">${scene.title}</option>`)
    .join('');
  document.querySelector<HTMLDivElement>('#app')!.innerHTML = `
<aside class="sidebar">
  <a class="brand" href="#budget">${logo}<span>OpenSheet<span class="brand-dot">.</span></span></a>
  <div class="workspace-tag"><span class="workspace-icon">O</span><div>Open workspace<small>Open-source workspace</small></div><span class="tag-caret">${icon(ChevronDown)}</span></div>
  <span class="nav-label">WORKSPACE</span>
  <button class="nav-item active" id="nav-sheet">${icon(Table2)} Spreadsheet <span class="nav-pill">1</span></button>
  <span class="nav-label space-top">DEVELOPER TOOLS</span>
  <button class="nav-item" id="nav-api">${icon(Braces)} API quick start</button>
  <button class="nav-item" id="nav-plugins">${icon(Braces)} Plugin example</button>
  <div class="sidebar-bottom"><div class="privacy-mark">${icon(ShieldCheck)} <span>Your data stays yours.</span></div><p>Files are processed on your device.<br>No account. No uploads.</p><div class="version"><span class="dot"></span> Open source <span>v0.1.1</span></div></div>
</aside>
<div class="page">
  <header class="topbar"><div class="breadcrumb">Workspace <span>/</span> <span id="page-name">Playground</span></div><div class="top-actions"><span class="local-chip"><span class="dot"></span> Runs locally</span><button class="icon-button" id="help" aria-label="Keyboard shortcuts">${icon(HelpCircle)}</button><span class="avatar">OS</span></div></header>
  <main>
    <div id="spreadsheet-page">
    <section class="workspace-heading" aria-label="Demo controls">
      <div class="scene-picker"><label for="scene-select">DEMO</label><select id="scene-select" aria-label="Choose demo">${sceneOptions}</select></div>
      <div class="scene-summary"><span class="eyebrow" id="scene-label"></span><h1 id="scene-title"></h1><p id="scene-description"></p></div>
      <button class="button secondary" id="reset-demo">${icon(RotateCcw)} Reset demo</button>
    </section>
    <div class="workbench" id="workbench">
      <section class="document" aria-label="Spreadsheet editor">
        <div class="document-bar"><div class="document-title"><span class="file-icon">${icon(FileSpreadsheet)}</span><div><input id="document-name" aria-label="Workbook name" value="Launch budget"><div class="doc-meta"><span class="dot"></span><span id="save-state">Example workbook · All changes stay local</span></div></div></div><div class="document-actions"><input id="file" type="file" accept=".opensheet,.xlsx,.csv,.tsv,.json" hidden><button class="button secondary" id="import">${icon(Upload)} <span>Import file</span></button><button class="button secondary" id="save-native">Save</button><button class="button secondary" id="save-json">Export JSON</button><button class="button primary" id="export-xlsx">${icon(Download)} <span>Export XLSX</span></button></div></div>
        <div class="utility-bar"><div class="search-wrap">${icon(Search)}<input id="find" placeholder="Find in this sheet…" aria-label="Find in sheet"><button id="find-next" aria-label="Find next">${icon(CornerDownLeft)}</button></div><div class="utility-controls"><select id="sheet-action" aria-label="Sheet actions"><option value="">Sheet actions</option><option value="sort-asc">Sort selection A → Z</option><option value="sort-desc">Sort selection Z → A</option><option value="delete-row">Delete selected rows</option><option value="delete-column">Delete selected columns</option><option value="hide-row">Hide selected row</option><option value="show-rows">Show all rows</option><option value="filter">Filter column by search</option><option value="clear-filter">Clear filter</option><option value="rename">Rename sheet</option><option value="remove">Delete sheet</option><option value="replace">Replace matches</option></select><label class="read-label"><input type="checkbox" id="read-only"> Read only</label><button id="accessible" class="text-button">Table view</button><select id="zoom" aria-label="Zoom"><option value="1">100%</option><option value=".75">75%</option><option value="1.25">125%</option><option value="1.5">150%</option></select></div></div>
        <div id="spreadsheet"></div>
      </section>
      <div class="splitter" id="splitter" role="separator" aria-label="Resize spreadsheet and results" aria-orientation="vertical" aria-valuemin="35" aria-valuemax="70" aria-valuenow="67" tabindex="0" title="Drag or use arrow keys to resize. Double-click to reset."><span aria-hidden="true"></span></div>
      <aside class="side-panel" aria-label="Demo results">
        <div class="panel-tabs" role="tablist" aria-label="Result panel">
          <button id="tab-insights" role="tab" aria-controls="panel-insights" aria-selected="true">Data view</button>
          <button id="tab-code" role="tab" aria-controls="panel-code" aria-selected="false">Code export</button>
        </div>
        <section class="scene-overview" id="panel-insights" role="tabpanel" aria-labelledby="tab-insights"><div id="scene-insights" aria-live="polite"></div><p class="scene-hint" id="scene-hint">Edit the sheet. Your changes update this view in real time.</p></section>
        <section class="source-card" id="panel-code" role="tabpanel" aria-labelledby="tab-code" hidden><div class="source-heading"><div><span class="small-icon">${icon(Code2)}</span><h2>From cells to code</h2><span class="live-badge">LIVE</span></div><button id="copy-code" class="text-button">${icon(Copy)} Copy code</button></div><div class="source-tabs" role="tablist" aria-label="Export format"><button class="selected" data-format="latex" role="tab" aria-selected="true">LaTeX</button><button data-format="markdown" role="tab" aria-selected="false">Markdown</button><button data-format="html" role="tab" aria-selected="false">HTML</button><button data-format="csv" role="tab" aria-selected="false">CSV</button><button data-format="tsv" role="tab" aria-selected="false">TSV</button><label><input id="export-selection" type="checkbox"> Selected cells only</label></div><pre><code id="source"></code></pre><div class="source-footer"><span id="export-note">Ready to paste into your next project.</span><button id="download-code" class="text-button">${icon(Download)} Download</button></div></section>
      </aside>
    </div>
    </div>
    <section class="developer-page" id="api-page" aria-labelledby="api-title" hidden><button class="text-button developer-back" id="api-back">← Back to spreadsheet</button><span class="eyebrow">DEVELOPER TOOLS</span><h1 id="api-title" tabindex="-1">API quick start</h1><p>Create a workbook, write values and formulas, and listen for changes.</p><pre><code id="api-example"></code></pre></section>
    <section class="developer-page" id="plugins-page" aria-labelledby="plugins-title" hidden><button class="text-button developer-back" id="plugins-back">← Back to spreadsheet</button><span class="eyebrow">DEVELOPER TOOLS</span><h1 id="plugins-title" tabindex="-1">Plugin example</h1><p>The Σ Sum toolbar button is a plugin. It reads the selected cells and displays their sum.</p><pre><code id="plugin-example"></code></pre></section>
  </main>
</div>
<dialog id="dialog"><div class="dialog-header"><h2 id="dialog-title"></h2><button id="dialog-close" aria-label="Close dialog">${icon(X)}</button></div><div id="dialog-content"></div></dialog><div id="notice" role="status" hidden></div>`;
  document.getElementById('api-example')!.textContent = apiExample;
  document.getElementById('plugin-example')!.textContent = pluginExample;
}

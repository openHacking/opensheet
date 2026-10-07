import {
  createOpenSheet,
  createWorkbook,
  address,
  definePlugin,
  type WorkbookSnapshot,
} from 'opensheet';
import { exportRange, escapeHTML, type Format } from '@opensheetjs/formats';
import { toSheetJS, type CompatibilityReport } from '@opensheetjs/adapter-sheetjs';
import 'opensheet/style.css';
import './style.css';
import { scenes, createScene, type SceneId } from './scenes';
import {
  createElement,
  Table2,
  BookOpen,
  Code2,
  Braces,
  ShieldCheck,
  HelpCircle,
  RotateCcw,
  FileSpreadsheet,
  Upload,
  Download,
  Search,
  CornerDownLeft,
  Copy,
  X,
  ChevronDown,
  type IconNode,
} from 'lucide';
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
  <button class="nav-item" id="nav-templates">${icon(BookOpen)} Sample workbooks</button>
  <span class="nav-label space-top">DEVELOPER TOOLS</span>
  <button class="nav-item" id="nav-code">${icon(Code2)} Table generator</button>
  <button class="nav-item" id="nav-api">${icon(Braces)} API quick start</button>
  <button class="nav-item" id="nav-plugins">${icon(Braces)} Plugin example</button>
  <div class="sidebar-bottom"><div class="privacy-mark">${icon(ShieldCheck)} <span>Your data stays yours.</span></div><p>Files are processed on your device.<br>No account. No uploads.</p><div class="version"><span class="dot"></span> Open source <span>v0.1.0</span></div></div>
</aside>
<div class="page">
  <header class="topbar"><div class="breadcrumb">Workspace <span>/</span> Playground</div><div class="top-actions"><span class="local-chip"><span class="dot"></span> Runs locally</span><button class="icon-button" id="help" aria-label="Keyboard shortcuts">${icon(HelpCircle)}</button><span class="avatar">OS</span></div></header>
  <main>
    <section class="workspace-heading" aria-label="Demo controls">
      <div class="scene-picker"><label for="scene-select">DEMO</label><select id="scene-select" aria-label="Choose demo">${sceneOptions}</select></div>
      <div class="scene-summary"><span class="eyebrow" id="scene-label"></span><h1 id="scene-title"></h1><p id="scene-description"></p></div>
      <button class="button secondary" id="reset-demo">${icon(RotateCcw)} Reset demo</button>
    </section>
    <div class="workbench" id="workbench">
      <section class="document" aria-label="Spreadsheet editor">
        <div class="document-bar"><div class="document-title"><span class="file-icon">${icon(FileSpreadsheet)}</span><div><input id="document-name" aria-label="Workbook name" value="Launch budget"><div class="doc-meta"><span class="dot"></span><span id="save-state">Example workbook · All changes stay local</span></div></div></div><div class="document-actions"><input id="file" type="file" accept=".xlsx,.csv,.tsv,.json" hidden><button class="button secondary" id="import">${icon(Upload)} <span>Import file</span></button><button class="button secondary" id="save-json">Save JSON</button><button class="button primary" id="export-xlsx">${icon(Download)} <span>Export XLSX</span></button></div></div>
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
  </main>
</div>
<dialog id="dialog"><div class="dialog-header"><h2 id="dialog-title"></h2><button id="dialog-close" aria-label="Close dialog">${icon(X)}</button></div><div id="dialog-content"></div></dialog><div id="notice" role="status" hidden></div>`;
const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
function notice(message: string) {
  $('notice').textContent = message;
  $('notice').hidden = false;
  setTimeout(() => ($('notice').hidden = true), 5000);
}
function modal(title: string, content: HTMLElement) {
  $('dialog-title').textContent = title;
  $('dialog-content').replaceChildren(content);
  $('dialog').scrollTop = 0;
  ($('dialog') as HTMLDialogElement).showModal();
}
$('dialog-close').onclick = () => ($('dialog') as HTMLDialogElement).close();
function textModal(title: string, text: string) {
  const pre = document.createElement('pre');
  pre.textContent = text;
  modal(title, pre);
}
const app = createOpenSheet({
  container: '#spreadsheet',
  onError: (e) => notice(e instanceof Error ? e.message : String(e)),
});
function sample(): WorkbookSnapshot {
  const b = createWorkbook({
    sheets: [
      { name: 'Launch budget', rows: 100, columns: 12 },
      { name: 'Notes', rows: 100, columns: 12 },
    ],
  });
  const s = b.getSheets()[0];
  b.transaction({ label: 'Example' }, () => {
    s.range('A1:G11').setValues([
      ['Item', 'Category', 'Owner', 'Quantity', 'Unit cost', 'Total', 'Status'],
      ['Brand identity', 'Design', 'Olivia', 1, 2400, null, 'Complete'],
      ['Landing page', 'Development', 'Alex', 1, 3800, null, 'In progress'],
      ['Product photography', 'Content', 'Sam', 2, 450, null, 'Complete'],
      ['Launch video', 'Content', 'Jordan', 1, 1800, null, 'In progress'],
      ['Social campaign', 'Marketing', 'Olivia', 4, 350, null, 'Planned'],
      ['Email platform', 'Software', 'Alex', 3, 49, null, 'Complete'],
      ['Press kit', 'Content', 'Sam', 1, 650, null, 'Planned'],
      ['Community event', 'Marketing', 'Jordan', 1, 1200, null, 'Planned'],
      ['Contingency', 'Operations', 'Alex', 1, 800, null, 'Reserved'],
      ['TOTAL', null, null, null, null, null, null],
    ]);
    s.range('F2:F10').setFormulas(Array.from({ length: 9 }, (_, i) => [`D${i + 2}*E${i + 2}`]));
    s.range('F11').setFormulas([['SUM(F2:F10)']]);
    s.range('A1:G1').setStyle({ bold: true, background: '#f0f5f1', color: '#476451' });
    s.range('E2:F11').setStyle({ numberFormat: '$#,##0.00' });
    s.range('A11:G11').setStyle({ bold: true, background: '#eef6f0' });
    s.range('G2:G10').setStyle({ color: '#248363', fontSize: 12 });
    s.setColumnWidth(0, 205);
    s.setColumnWidth(1, 142);
    s.setColumnWidth(2, 112);
    s.setColumnWidth(3, 92);
    s.setColumnWidth(4, 118);
    s.setColumnWidth(5, 130);
    s.setColumnWidth(6, 145);
    s.setRowHeight(0, 36);
    s.setFreeze(1);
    b.getSheets()[1]
      .range('A1:B5')
      .setValues([
        ['OpenSheet', 'Quick notes'],
        ['Edit', 'Double-click a cell or press Enter.'],
        ['Formula', '=SUM(F2:F10)'],
        ['Privacy', 'All processing happens on this device.'],
        ['Compatibility', 'XLSX exports preserve supported data, not every Excel feature.'],
      ]);
    b.getSheets()[1].setColumnWidth(1, 540);
  });
  const json = b.toJSON();
  b.dispose();
  return json;
}
await app.load(sample());
const sumPlugin = definePlugin({
  id: 'example.selection-sum',
  version: '0.1.0',
  apiVersion: '^0.1.0',
  capabilities: ['selection.read', 'workbook.read', 'ui.toolbar'],
  setup(ctx) {
    ctx.ui.toolbar.add({
      id: 'example.selection-sum.run',
      label: 'Σ Sum',
      run() {
        const s = ctx.selection.get();
        if (s) {
          const values = ctx.workbook.readRange(s).values.flat();
          const total = values.reduce<number>((sum, v) => sum + (typeof v === 'number' ? v : 0), 0);
          ctx.ui.notify(`Selected numbers sum to ${total.toLocaleString()}`);
        }
      },
    });
  },
});
app.use(sumPlugin);
let currentScene: SceneId = 'budget';
let dirty = false;
let importedWorkbook = false;
const currency = (value: number) =>
  new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0,
  }).format(value);
const numeric = (value: unknown) =>
  typeof value === 'number' && Number.isFinite(value) ? value : 0;
function metric(label: string, value: string, detail: string) {
  return `<div class="metric"><span>${label}</span><strong>${escapeHTML(value)}</strong><small>${detail}</small></div>`;
}
function renderInsights() {
  const sheet = app.getWorkbook().getSheets()[0];
  const rows = sheet.getUsedRange().getValues();
  let html = '';
  if (importedWorkbook) {
    $('scene-insights').innerHTML =
      `<div class="metrics">${metric('Workbook', 'Imported', 'Save JSON to keep your edits')}${metric('Sheets', String(app.getWorkbook().getSheets().length), 'Imported data replaces the sample')}${metric('Rows in first sheet', String(rows.length), 'Sample-specific visualizations are paused')}</div>`;
    return;
  }
  if (currentScene === 'budget') {
    const total = rows
      .slice(1)
      .reduce((sum, row) => sum + (row[0] === 'TOTAL' ? 0 : numeric(row[5])), 0);
    const ceiling = 18000;
    html = `<div class="metrics">${metric('Planned spend', currency(total), 'Calculated from the workbook')}${metric('Launch reserve', currency(ceiling - total), 'Against an $18,000 demo budget')}${metric('Budget allocated', `${Math.round((total / ceiling) * 100)}%`, 'Edit quantity or unit cost below')}</div><div class="budget-track" role="img" aria-label="Budget allocated ${Math.round((total / ceiling) * 100)} percent"><span style="width:${Math.max(0, Math.min(100, (total / ceiling) * 100))}%"></span></div>`;
  } else if (currentScene === 'sales') {
    const data = rows.slice(1).filter((row) => row[0] !== 'TOTAL');
    const total = data.reduce((sum, row) => sum + numeric(row[1]), 0);
    const target = data.reduce((sum, row) => sum + numeric(row[2]), 0);
    const max = Math.max(1, ...data.map((row) => Math.max(numeric(row[1]), numeric(row[2]))));
    html = `<div class="metrics">${metric('Revenue', currency(total), 'All months in this workbook')}${metric('Target', currency(target), 'Your editable sales goals')}${metric('Attainment', target ? `${Math.round((total / target) * 100)}%` : '—', 'Revenue ÷ target')}</div><div class="chart" role="img" aria-label="Monthly revenue and target chart"><div class="chart-key"><span>● Revenue</span><span>○ Target</span></div><div class="chart-bars">${data
      .slice(0, 24)
      .map(
        (row) =>
          `<div class="chart-column"><div class="bar-pair"><span class="revenue-bar" style="height:${Math.max(0, (numeric(row[1]) / max) * 100)}%" title="${escapeHTML(currency(numeric(row[1])))}"></span><span class="target-bar" style="height:${Math.max(0, (numeric(row[2]) / max) * 100)}%"></span></div><small>${escapeHTML(String(row[0] ?? ''))}</small></div>`,
      )
      .join('')}</div></div>`;
  } else if (currentScene === 'planner') {
    const data = rows
      .slice(1)
      .filter((row) => row[0] != null)
      .slice(0, 30);
    const dates = data.map((row) =>
      typeof row[2] === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(row[2])
        ? Date.parse(`${row[2]}T00:00:00Z`) / 86400000
        : NaN,
    );
    const starts = dates.filter(Number.isFinite);
    const first = starts.length ? Math.min(...starts) : 0;
    const span = Math.max(
      30,
      ...data.map((row, i) =>
        Number.isFinite(dates[i]) ? dates[i] - first + Math.max(0, numeric(row[3])) : 0,
      ),
    );
    const progress = data.length
      ? data.reduce((sum, row) => sum + Math.max(0, Math.min(1, numeric(row[4]))), 0) / data.length
      : 0;
    html = `<div class="metrics">${metric('Tasks', String(data.length), 'A shared plan for the next launch')}${metric('Overall progress', `${Math.round(progress * 100)}%`, 'Mean completion across tasks')}${metric('Schedule', `${Math.ceil(span)} days`, 'Start dates use YYYY-MM-DD')}</div><div class="timeline" role="img" aria-label="Project task timeline">${data.map((row, i) => `<div class="timeline-row"><span>${escapeHTML(String(row[0]))}</span><div class="timeline-track">${Number.isFinite(dates[i]) ? `<div class="timeline-bar" style="left:${((dates[i] - first) / span) * 100}%;width:${(Math.max(0, numeric(row[3])) / span) * 100}%"><span style="width:${Math.max(0, Math.min(1, numeric(row[4]))) * 100}%"></span></div>` : '<small>Enter a valid start date</small>'}</div><small>${Math.round(Math.max(0, Math.min(1, numeric(row[4]))) * 100)}%</small></div>`).join('')}</div>`;
  } else {
    html = `<div class="metrics">${metric('Five formats', 'One source', 'LaTeX · Markdown · HTML · CSV · TSV')}${metric('Rows', String(Math.max(0, rows.length - 1)), 'Edit your table below')}${metric('Workflow', 'Edit → export', 'Copy or download generated code')}</div>`;
  }
  $('scene-insights').innerHTML = html;
}
type Panel = 'insights' | 'code';
function showPanel(panel: Panel) {
  for (const name of ['insights', 'code'] as const) {
    const active = name === panel;
    $(`panel-${name}`).hidden = !active;
    $(`tab-${name}`).setAttribute('aria-selected', String(active));
    $<HTMLButtonElement>(`tab-${name}`).tabIndex = active ? 0 : -1;
  }
}
for (const name of ['insights', 'code'] as const) {
  $(`tab-${name}`).onclick = () => showPanel(name);
  $(`tab-${name}`).onkeydown = (event) => {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
    event.preventDefault();
    const next = name === 'insights' ? 'code' : 'insights';
    showPanel(next);
    $(`tab-${next}`).focus();
  };
}
const workbench = $('workbench');
const splitter = $('splitter');
const defaultSplitFraction = 0.67;
let splitFraction = defaultSplitFraction;
function resizeWorkbench(fraction = splitFraction) {
  if (getComputedStyle(workbench).display !== 'grid') return;
  const available = workbench.clientWidth - splitter.offsetWidth;
  if (available <= 0) return;
  const min = Math.min(520, Math.max(0, available - 320));
  const max = Math.max(min, available - 320);
  const left = Math.max(min, Math.min(max, available * fraction));
  splitFraction = left / available;
  workbench.style.setProperty('--sheet-width', `${left}px`);
  splitter.setAttribute('aria-valuemin', String(Math.round((min / available) * 100)));
  splitter.setAttribute('aria-valuemax', String(Math.round((max / available) * 100)));
  splitter.setAttribute('aria-valuenow', String(Math.round(splitFraction * 100)));
}
new ResizeObserver(() => resizeWorkbench()).observe(workbench);
splitter.addEventListener('pointerdown', (event) => {
  if (event.button !== 0) return;
  splitter.setPointerCapture(event.pointerId);
  splitter.classList.add('dragging');
  event.preventDefault();
});
splitter.addEventListener('pointermove', (event) => {
  if (!splitter.hasPointerCapture(event.pointerId)) return;
  const available = workbench.clientWidth - splitter.offsetWidth;
  resizeWorkbench((event.clientX - workbench.getBoundingClientRect().left) / available);
});
splitter.addEventListener('pointerup', (event) => {
  if (splitter.hasPointerCapture(event.pointerId)) splitter.releasePointerCapture(event.pointerId);
  splitter.classList.remove('dragging');
});
splitter.addEventListener('pointercancel', () => splitter.classList.remove('dragging'));
splitter.addEventListener('dblclick', () => resizeWorkbench(defaultSplitFraction));
splitter.addEventListener('keydown', (event) => {
  const available = workbench.clientWidth - splitter.offsetWidth;
  if (event.key === 'ArrowLeft') resizeWorkbench(splitFraction - 24 / available);
  else if (event.key === 'ArrowRight') resizeWorkbench(splitFraction + 24 / available);
  else if (event.key === 'Home') resizeWorkbench(0);
  else if (event.key === 'End') resizeWorkbench(1);
  else return;
  event.preventDefault();
});
async function loadScene(id: SceneId) {
  importWorker?.terminate();
  importWorker = undefined;
  await app.load(id === 'budget' ? sample() : createScene(id));
  currentScene = id;
  dirty = false;
  importedWorkbook = false;
  foundIndex = -1;
  $<HTMLInputElement>('read-only').checked = false;
  app.setMode('edit');
  $<HTMLInputElement>('document-name').value = id === 'budget' ? 'Launch budget' : scenes[id].title;
  $('save-state').textContent = 'Example workbook · All changes stay local';
  $('scene-title').textContent = scenes[id].title;
  $('scene-label').textContent = scenes[id].label;
  $('scene-description').textContent = scenes[id].description;
  document.title = `${scenes[id].title} · OpenSheet`;
  $<HTMLSelectElement>('scene-select').value = id;
  showPanel(id === 'code' ? 'code' : 'insights');
  refresh();
}
let format: Format = 'latex',
  sourceText = '',
  refreshTimer: ReturnType<typeof setTimeout> | undefined;
function refresh() {
  renderInsights();
  clearTimeout(refreshTimer);
  refreshTimer = setTimeout(() => {
    try {
      const s = app.selection.get()!;
      const config = {
        sheetId: s.sheetId,
        format,
        ...($<HTMLInputElement>('export-selection').checked ? { range: s } : {}),
      };
      const result = exportRange(app.getWorkbook().toJSON(), config);
      sourceText = result.text;
      $('source').textContent =
        sourceText.length > 30000
          ? sourceText.slice(0, 30000) + '\n… Preview truncated; download includes the full result.'
          : sourceText;
      $('export-note').textContent =
        result.report.join(' ') ||
        `${format === 'latex' ? 'HTML grid preview; actual TeX layout may differ.' : 'Ready to paste into your next project.'}`;
    } catch (e) {
      $('source').textContent = e instanceof Error ? e.message : String(e);
      sourceText = '';
    }
  }, 120);
}
app.on('workbook:committed', () => {
  dirty = true;
  $('save-state').textContent = 'Unsaved local changes · Download to keep your work';
  refresh();
});
app.on('selection:changed', refresh);
refresh();
for (const b of document.querySelectorAll<HTMLButtonElement>('[data-format]'))
  b.onclick = () => {
    format = b.dataset.format as Format;
    for (const x of document.querySelectorAll('[data-format]')) {
      x.classList.toggle('selected', x === b);
      x.setAttribute('aria-selected', String(x === b));
    }
    refresh();
  };
$<HTMLInputElement>('export-selection').onchange = refresh;
function download(data: BlobPart, name: string, type = 'text/plain') {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function name() {
  return $<HTMLInputElement>('document-name').value.trim() || 'workbook';
}
$('download-code').onclick = () =>
  download(
    sourceText,
    `${name()}.${format === 'latex' ? 'tex' : format === 'markdown' ? 'md' : format}`,
  );
$('copy-code').onclick = async () => {
  try {
    await navigator.clipboard.writeText(sourceText);
    notice('Code copied to clipboard');
  } catch {
    textModal('Copy this code', sourceText);
  }
};
$('save-json').onclick = () => {
  download(
    JSON.stringify(app.getWorkbook().toJSON(), null, 2),
    `${name()}.opensheet.json`,
    'application/json',
  );
  dirty = false;
  $('save-state').textContent = 'JSON downloaded · Continue editing locally';
};
function showReport(report: CompatibilityReport, proceed?: () => void) {
  const box = document.createElement('div');
  const p = document.createElement('p');
  p.textContent = `${report.summary.exact} supported entries · ${report.summary.approximated} approximations · ${report.summary.dropped} omitted features`;
  box.append(p);
  const list = document.createElement('ul');
  for (const issue of report.issues) {
    const item = document.createElement('li');
    item.textContent = `${issue.message}${issue.count ? ` (${issue.count} occurrences)` : ''}`;
    list.append(item);
  }
  box.append(list);
  if (proceed) {
    const b = document.createElement('button');
    b.className = 'button primary';
    b.textContent = 'Download converted XLSX';
    b.onclick = () => {
      proceed();
      ($('dialog') as HTMLDialogElement).close();
    };
    box.append(b);
  }
  modal('Compatibility report', box);
}
$('export-xlsx').onclick = () => {
  try {
    const result = toSheetJS(app.getWorkbook().toJSON());
    const save = async () => {
      try {
        const XLSX = await import('xlsx');
        const bytes = XLSX.write(result.workbook, { type: 'array', bookType: 'xlsx' });
        download(
          bytes,
          `${name()}.xlsx`,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        );
      } catch (e) {
        notice(String(e));
      }
    };
    if (result.report.issues.length) showReport(result.report, save);
    else save();
  } catch (e) {
    notice(String(e));
  }
};
$('import').onclick = () => $<HTMLInputElement>('file').click();
let importWorker: Worker | undefined;
$<HTMLInputElement>('file').onchange = async () => {
  const file = $<HTMLInputElement>('file').files?.[0];
  if (!file) return;
  if (file.size > 20 * 1024 * 1024) {
    notice('Files must be smaller than 20 MiB');
    return;
  }
  importWorker?.terminate();
  const worker = new Worker(new URL('./file.worker.ts', import.meta.url), { type: 'module' });
  importWorker = worker;
  const b = $<HTMLButtonElement>('import');
  b.disabled = true;
  b.textContent = 'Reading…';
  let timer: ReturnType<typeof setTimeout>;
  const stop = () => {
    worker.terminate();
    clearTimeout(timer);
    b.disabled = false;
    b.textContent = '↑ Import file';
    $<HTMLInputElement>('file').value = '';
  };
  timer = setTimeout(() => {
    stop();
    notice('Import timed out. Your current workbook is unchanged.');
  }, 20000);
  worker.onerror = () => {
    stop();
    notice('Could not parse this file.');
  };
  worker.onmessage = async (
    event: MessageEvent<{
      error?: string;
      snapshot: WorkbookSnapshot;
      report?: CompatibilityReport;
    }>,
  ) => {
    stop();
    if (event.data.error) {
      notice(event.data.error);
      return;
    }
    try {
      await app.load(event.data.snapshot);
      dirty = true;
      importedWorkbook = true;
      $<HTMLInputElement>('document-name').value = file.name.replace(/\.(xlsx|csv|tsv|json)$/i, '');
      $('save-state').textContent = 'Imported locally · Original file unchanged';
      refresh();
      if (event.data.report?.issues.length) showReport(event.data.report);
    } catch (e) {
      notice(String(e));
    }
  };
  try {
    const buffer = await file.arrayBuffer();
    worker.postMessage({ buffer, name: file.name }, [buffer]);
  } catch (e) {
    stop();
    notice(String(e));
  }
};
$<HTMLInputElement>('read-only').onchange = (e) =>
  app.setMode((e.target as HTMLInputElement).checked ? 'read' : 'edit');
$<HTMLSelectElement>('zoom').onchange = (e) =>
  app.getGrid()?.setZoom(Number((e.target as HTMLSelectElement).value));
let foundIndex = -1;
function findNext() {
  const query = $<HTMLInputElement>('find').value.toLowerCase();
  if (!query) return;
  const s = app.selection.get()!,
    sheet = app.getWorkbook().getSheetById(s.sheetId)!,
    r = sheet.getUsedRange().bounds,
    matches: Array<[number, number]> = [];
  for (let row = 0; row < r.endRow; row++)
    for (let c = 0; c < r.endColumn; c++)
      if (app.getWorkbook().display(s.sheetId, row, c).toLowerCase().includes(query))
        matches.push([row, c]);
  if (!matches.length) {
    notice('No matching cells');
    return;
  }
  foundIndex = (foundIndex + 1) % matches.length;
  const [row, col] = matches[foundIndex];
  app.selection.set({
    sheetId: s.sheetId,
    startRow: row,
    endRow: row + 1,
    startColumn: col,
    endColumn: col + 1,
  });
  notice(`${foundIndex + 1} of ${matches.length} matches`);
}
$('find-next').onclick = findNext;
$('find').onkeydown = (e) => {
  if (e.key === 'Enter') findNext();
};
$('find').oninput = () => {
  foundIndex = -1;
};
$<HTMLSelectElement>('sheet-action').onchange = (e) => {
  const select = e.target as HTMLSelectElement,
    value = select.value;
  select.value = '';
  try {
    const book = app.getWorkbook(),
      s = app.selection.get()!,
      sh = book.getSheetById(s.sheetId)!;
    switch (value) {
      case 'sort-asc':
      case 'sort-desc':
        sh.sort(s, s.startColumn, value === 'sort-asc' ? 'asc' : 'desc');
        break;
      case 'delete-row':
        sh.deleteRows(s.startRow, s.endRow - s.startRow);
        break;
      case 'delete-column':
        sh.deleteColumns(s.startColumn, s.endColumn - s.startColumn);
        break;
      case 'hide-row':
        sh.setRowHidden(s.startRow, true);
        break;
      case 'show-rows':
        book.transaction({}, () => {
          const data = book.sheetData(sh.id);
          data.rowOrder.forEach((id, i) => {
            if (data.rows[id]?.hidden) sh.setRowHidden(i, false);
          });
        });
        break;
      case 'filter':
        sh.setFilter({ column: s.startColumn, query: $<HTMLInputElement>('find').value });
        break;
      case 'clear-filter':
        sh.setFilter(null);
        break;
      case 'rename': {
        const next = prompt('Sheet name', sh.name);
        if (next) book.renameSheet(sh.id, next);
        break;
      }
      case 'remove':
        if (confirm(`Delete sheet “${sh.name}”? You can undo this.`)) book.removeSheet(sh.id);
        break;
      case 'replace': {
        const query = $<HTMLInputElement>('find').value;
        if (!query) {
          notice('Enter search text first');
          return;
        }
        const replacement = prompt(`Replace “${query}” with:`);
        if (replacement === null) return;
        book.transaction({}, () => {
          const data = book.sheetData(sh.id);
          for (const c of Object.values(data.cells))
            if (c.input.type === 'string' && c.input.value.includes(query)) {
              const row = data.rowOrder.indexOf(c.rowId),
                col = data.columnOrder.indexOf(c.columnId);
              sh.range(address(row, col)).setValues([
                [c.input.value.replaceAll(query, replacement)],
              ]);
            }
        });
        break;
      }
    }
  } catch (error) {
    notice(error instanceof Error ? error.message : String(error));
  }
};
$('accessible').onclick = () => {
  const s = app.selection.get()!,
    sheet = app.getWorkbook().getSheetById(s.sheetId)!,
    range = sheet.getUsedRange(),
    box = document.createElement('div'),
    table = document.createElement('table'),
    caption = document.createElement('caption');
  caption.textContent = `${sheet.name} — first 200 rows (read-only accessible view)`;
  table.append(caption);
  for (const row of range.getDisplayValues().slice(0, 200)) {
    const tr = document.createElement('tr');
    for (const v of row) {
      const td = document.createElement('td');
      td.textContent = v;
      tr.append(td);
    }
    table.append(tr);
  }
  box.append(table);
  modal('Accessible table view', box);
};
const api = `import { createOpenSheet } from 'opensheet';\nimport 'opensheet/style.css';\n\nconst app = createOpenSheet({ container: '#sheet' });\nconst book = app.createWorkbook();\nconst sheet = book.getSheets()[0];\n\nsheet.range('A1:B2').setValues([\n  ['Product', 'Revenue'], ['OpenSheet', 1200]\n]);\nsheet.range('B3').setFormulas([['SUM(B2:B2)']]);\napp.on('workbook:committed', ({ revision }) => {\n  console.log(revision, book.toJSON());\n});\n\n// On unmount\napp.dispose();`;
$('nav-api').onclick = () => textModal('API quick start', api);
$('nav-plugins').onclick = () =>
  textModal(
    'The Σ Sum button is a plugin',
    `definePlugin({\n  id: 'example.selection-sum',\n  version: '0.1.0',\n  apiVersion: '^0.1.0',\n  capabilities: ['selection.read', 'workbook.read', 'ui.toolbar'],\n  setup(ctx) {\n    ctx.ui.toolbar.add({\n      id: 'sum', label: 'Σ Sum',\n      run() {\n        const selection = ctx.selection.get();\n        const rows = ctx.workbook.readRange(selection).values;\n        const sum = rows.flat().reduce(\n          (total, n) => total + (typeof n === 'number' ? n : 0), 0\n        );\n        ctx.ui.notify(String(sum));\n      }\n    });\n  }\n});`,
  );
$('help').onclick = () =>
  textModal(
    'Keyboard shortcuts',
    'Enter / F2     Edit selected cell\nTab / Shift+Tab     Next / previous cell\nArrow keys     Move selection\nShift + arrows     Extend selection\nCtrl/Cmd + C / V     Copy / paste\nCtrl/Cmd + Z     Undo\nCtrl/Cmd + Shift + Z     Redo\nDelete / Backspace     Clear values\nEscape     Cancel editing\n\nPaste values starting with = as formulas. Use an apostrophe to enter literal formula text.\n\n† after a value indicates an imported formula cache that has not been recalculated.',
  );
$('nav-code').onclick = () => showPanel('code');
$('nav-sheet').onclick = () =>
  document.querySelector<HTMLElement>('#spreadsheet .os-grid')?.focus();
$('nav-templates').onclick = () => $<HTMLSelectElement>('scene-select').focus();
$<HTMLSelectElement>('scene-select').onchange = (event) => {
  const next = (event.target as HTMLSelectElement).value as SceneId;
  if (next === currentScene) return;
  location.hash = next;
};
$('reset-demo').onclick = async () => {
  if (dirty && !confirm('Reset this demo? Download JSON first to keep your changes.')) return;
  await loadScene(currentScene);
};
// The playground exposes the actual public API for learning and browser contract tests.
Object.assign(window, { opensheet: app });
window.addEventListener('beforeunload', (event) => {
  if (dirty) {
    event.preventDefault();
    event.returnValue = '';
  }
});
window.addEventListener('pagehide', () => {
  importWorker?.terminate();
  app.dispose();
});

function requestedScene(): SceneId {
  const id = location.hash.slice(1);
  return Object.hasOwn(scenes, id) ? (id as SceneId) : 'budget';
}
let routeBusy = false;
window.addEventListener('hashchange', async () => {
  const next = requestedScene();
  if (routeBusy || next === currentScene) {
    $<HTMLSelectElement>('scene-select').value = currentScene;
    return;
  }
  if (dirty && !confirm('Switch demos? Download JSON first to keep your changes.')) {
    history.replaceState(null, '', `#${currentScene}`);
    $<HTMLSelectElement>('scene-select').value = currentScene;
    return;
  }
  routeBusy = true;
  try {
    await loadScene(next);
  } finally {
    routeBusy = false;
  }
});
await loadScene(requestedScene());

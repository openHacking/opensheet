import { escapeHTML } from '@opensheetjs/formats';
import type { OpenSheet } from 'opensheet';
import { $ } from './dom.js';
import type { PlaygroundState } from './state.js';
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
export function renderInsights(app: OpenSheet, state: PlaygroundState) {
  const sheet = app.getWorkbook().getSheets()[0];
  const rows = sheet.getUsedRange().getValues();
  let html = '';
  if (state.importedWorkbook) {
    $('scene-insights').innerHTML =
      `<div class="metrics">${metric('Workbook', 'Imported', 'Save JSON to keep your edits')}${metric('Sheets', String(app.getWorkbook().getSheets().length), 'Imported data replaces the sample')}${metric('Rows in first sheet', String(rows.length), 'Sample-specific visualizations are paused')}</div>`;
    return;
  }
  if (state.currentScene === 'budget') {
    const total = rows
      .slice(1)
      .reduce((sum, row) => sum + (row[0] === 'TOTAL' ? 0 : numeric(row[5])), 0);
    const ceiling = 18000;
    html = `<div class="metrics">${metric('Planned spend', currency(total), 'Calculated from the workbook')}${metric('Launch reserve', currency(ceiling - total), 'Against an $18,000 demo budget')}${metric('Budget allocated', `${Math.round((total / ceiling) * 100)}%`, 'Edit quantity or unit cost below')}</div><div class="budget-track" role="img" aria-label="Budget allocated ${Math.round((total / ceiling) * 100)} percent"><span style="width:${Math.max(0, Math.min(100, (total / ceiling) * 100))}%"></span></div>`;
  } else if (state.currentScene === 'sales') {
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
  } else if (state.currentScene === 'planner') {
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

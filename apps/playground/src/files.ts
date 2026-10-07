import { toSheetJS, type CompatibilityReport } from '@opensheetjs/adapter-sheetjs';
import type { OpenSheet, WorkbookSnapshot } from 'opensheet';
import { createBindings } from './bindings.js';
import { $ } from './dom.js';
import type { Feedback } from './feedback.js';
import type { PlaygroundState } from './state.js';
export function createFileActions(
  app: OpenSheet,
  state: PlaygroundState,
  { notice, modal }: Pick<Feedback, 'notice' | 'modal'>,
  refresh: () => void,
) {
  const bindings = createBindings();
  let disposed = false;
  const importButton = $<HTMLButtonElement>('import');
  const importLabel = importButton.innerHTML;
  const downloads = new Map<string, ReturnType<typeof setTimeout>>();
  function download(data: BlobPart, name: string, type = 'text/plain') {
    if (disposed) return;
    const url = URL.createObjectURL(new Blob([data], { type }));
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
    const timer = setTimeout(() => {
      URL.revokeObjectURL(url);
      downloads.delete(url);
    }, 1000);
    downloads.set(url, timer);
  }
  function name() {
    return $<HTMLInputElement>('document-name').value.trim() || 'workbook';
  }
  bindings.on($('save-json'), 'click', () => {
    download(
      JSON.stringify(app.getWorkbook().toJSON(), null, 2),
      `${name()}.opensheet.json`,
      'application/json',
    );
    state.dirty = false;
    $('save-state').textContent = 'JSON downloaded · Continue editing locally';
  });
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
      bindings.on(b, 'click', () => {
        proceed();
        ($('dialog') as HTMLDialogElement).close();
      });
      box.append(b);
    }
    modal('Compatibility report', box);
  }
  bindings.on($('export-xlsx'), 'click', () => {
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
  });
  bindings.on($('import'), 'click', () => $<HTMLInputElement>('file').click());
  let importWorker: Worker | undefined;
  let importTimer: ReturnType<typeof setTimeout> | undefined;
  let generation = 0;
  function cancelImport() {
    generation++;
    importWorker?.terminate();
    importWorker = undefined;
    clearTimeout(importTimer);
    importTimer = undefined;
    const b = $<HTMLButtonElement>('import');
    b.disabled = false;
    b.innerHTML = importLabel;
    $<HTMLInputElement>('file').value = '';
  }
  bindings.on($<HTMLInputElement>('file'), 'change', async () => {
    const file = $<HTMLInputElement>('file').files?.[0];
    if (!file) return;
    if (file.size > 20 * 1024 * 1024) {
      notice('Files must be smaller than 20 MiB');
      return;
    }
    cancelImport();
    const request = generation;
    const worker = new Worker(new URL('./file.worker.ts', import.meta.url), { type: 'module' });
    importWorker = worker;
    const active = () => request === generation && importWorker === worker;
    const b = $<HTMLButtonElement>('import');
    b.disabled = true;
    b.textContent = 'Reading…';
    const stop = () => {
      if (!active()) return;
      worker.onmessage = null;
      worker.onerror = null;
      worker.terminate();
      clearTimeout(importTimer);
      importTimer = undefined;
      if (importWorker === worker) importWorker = undefined;
      b.disabled = false;
      b.innerHTML = importLabel;
      $<HTMLInputElement>('file').value = '';
    };
    importTimer = setTimeout(() => {
      stop();
      notice('Import timed out. Your current workbook is unchanged.');
    }, 20000);
    worker.onerror = () => {
      if (!active()) return;
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
      if (!active()) return;
      stop();
      if (event.data.error) {
        notice(event.data.error);
        return;
      }
      try {
        await app.load(event.data.snapshot);
        if (request !== generation || disposed) return;
        state.dirty = true;
        state.importedWorkbook = true;
        $<HTMLInputElement>('document-name').value = file.name.replace(
          /\.(xlsx|csv|tsv|json)$/i,
          '',
        );
        $('save-state').textContent = 'Imported locally · Original file unchanged';
        refresh();
        if (event.data.report?.issues.length) showReport(event.data.report);
      } catch (e) {
        notice(String(e));
      }
    };
    try {
      const buffer = await file.arrayBuffer();
      if (!active()) return;
      worker.postMessage({ buffer, name: file.name }, [buffer]);
    } catch (e) {
      if (!active()) return;
      stop();
      notice(String(e));
    }
  });
  return {
    download,
    cancelImport,
    dispose() {
      disposed = true;
      cancelImport();
      bindings.dispose();
      for (const [url, timer] of downloads) {
        clearTimeout(timer);
        URL.revokeObjectURL(url);
      }
      downloads.clear();
    },
  };
}

import {
  unsupportedXlsxSheetName,
  XLSX_SHEET_NAME_MESSAGE,
  type CompatibilityReport,
} from '@opensheetjs/adapter-sheetjs';
import type { OpenSheet, WorkbookFile } from 'opensheet';
import { createBindings } from './bindings.js';
import { $ } from './dom.js';
import type { Feedback } from './feedback.js';
import type { PlaygroundState } from './state.js';
import { refreshPersistence } from './persistence.js';
const MAX_FOREIGN_FILE_BYTES = 20 * 1024 * 1024;
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
  let exportWorker: Worker | undefined;
  let exportTimer: ReturnType<typeof setTimeout> | undefined;
  const exportButton = $<HTMLButtonElement>('export-xlsx');
  const exportLabel = exportButton.innerHTML;
  function stopExport() {
    clearTimeout(exportTimer);
    exportTimer = undefined;
    exportWorker?.terminate();
    exportWorker = undefined;
    exportButton.disabled = false;
    exportButton.innerHTML = exportLabel;
  }
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
  function captureExport() {
    const book = app.getWorkbook();
    const revision = book.revision;
    const check = () => {
      if (app.hasDraft || book.persistenceState.pending)
        throw new Error('Wait for saving to finish');
      if (app.getWorkbook() !== book || book.revision !== revision)
        throw new Error('Workbook changed; export again');
    };
    check();
    return { book, check };
  }
  bindings.on($('save-native'), 'click', async () => {
    try {
      const { book, check } = captureExport();
      const parts: BlobPart[] = [];
      for await (const chunk of book.streamBinary()) parts.push(chunk as BlobPart);
      check();
      download(new Blob(parts), `${name()}.opensheet`, 'application/octet-stream');
      state.dirty = false;
      notice('Workbook downloaded');
    } catch (error) {
      notice(String(error));
    }
  });
  bindings.on($('save-json'), 'click', async () => {
    try {
      const { book, check } = captureExport();
      const parts: string[] = [];
      for await (const chunk of book.streamJSON()) parts.push(chunk);
      check();

      download(
        new Blob(parts, { type: 'application/json' }),
        `${name()}.opensheet.json`,
        'application/json',
      );
      state.dirty = false;
      notice('JSON downloaded');
    } catch (error) {
      notice(String(error));
    }
  });
  function showReport(
    report: CompatibilityReport,
    proceed?: () => void,
    label = 'Download converted XLSX',
    warning?: string,
    sheets: ReadonlyArray<{ id: string; name: string }> = [],
  ) {
    if ($<HTMLDialogElement>('dialog').open) {
      notice('Close the current dialog and try again.');
      return;
    }
    const sheetNames = new Map(sheets.map((sheet) => [sheet.id, sheet.name]));
    const box = document.createElement('div');
    const p = document.createElement('p');
    p.textContent = `${report.summary.exact} supported entries · ${report.summary.approximated} approximations · ${report.summary.dropped} omitted features · ${report.summary.blocked} blocked features`;
    box.append(p);
    if (warning) {
      const warningText = document.createElement('p');
      warningText.textContent = warning;
      box.append(warningText);
    }
    if (report.omittedIssueOccurrences) {
      const notice = document.createElement('p');
      notice.textContent = `Details limited to 200 groups; ${report.omittedIssueOccurrences} additional occurrences are included in the totals above.`;
      box.append(notice);
    }
    const list = document.createElement('ul');
    for (const issue of report.issues) {
      const item = document.createElement('li');
      item.textContent = `${issue.sheetId ? `${sheetNames.get(issue.sheetId) ?? issue.sheetId}: ` : ''}${issue.address ? `${issue.address}: ` : ''}${issue.message}${issue.count && issue.count > 1 ? ` (${issue.count} occurrences${issue.address ? '; first location shown' : ''})` : ''}`;
      list.append(item);
    }
    box.append(list);
    if (proceed) {
      const b = document.createElement('button');
      b.tabIndex = 0;
      b.className = 'button primary';
      b.textContent = label;
      const dialog = $<HTMLDialogElement>('dialog');
      dialog.addEventListener(
        'close',
        () => {
          proceed = undefined;
        },
        { once: true, signal: bindings.signal },
      );
      bindings.on(b, 'click', () => {
        if (!b.isConnected || !dialog.open) return;
        const action = proceed;
        proceed = undefined;
        dialog.close();
        action?.();
      });
      box.append(b);
    }
    modal('Compatibility report', box);
  }
  bindings.on(exportButton, 'click', async () => {
    if (exportButton.disabled) return;
    exportButton.disabled = true;
    exportButton.textContent = 'Preparing…';
    try {
      const { book, check } = captureExport();
      if (book.getSheets().some((sheet) => unsupportedXlsxSheetName(sheet.name)))
        throw new Error(XLSX_SHEET_NAME_MESSAGE);
      const file = await book.toJSON();
      if (disposed) return;
      check();
      const worker = new Worker(new URL('./xlsx-export.worker.ts', import.meta.url), {
        type: 'module',
      });
      exportWorker = worker;
      const active = () => !disposed && exportWorker === worker;
      exportTimer = setTimeout(() => {
        if (!active()) return;
        stopExport();
        notice('XLSX export timed out.');
      }, 120000);
      const fail = (message: string) => {
        if (!active()) return;
        stopExport();
        notice(message);
      };
      worker.onerror = () => fail('Could not prepare XLSX.');
      worker.addEventListener('messageerror', () => fail('Could not receive XLSX.'));
      worker.onmessage = (
        event: MessageEvent<{ buffer?: ArrayBuffer; report?: CompatibilityReport; error?: string }>,
      ) => {
        if (!active()) return;
        if (event.data.error && event.data.report) {
          const { error, report } = event.data;
          stopExport();
          try {
            check();
            showReport(report, undefined, undefined, error, file.sheets);
          } catch (e) {
            notice(String(e));
          }
          return;
        }
        if (event.data.error || !event.data.buffer || !event.data.report) {
          fail(event.data.error || 'Could not prepare XLSX.');
          return;
        }
        const { buffer, report } = event.data;
        stopExport();
        try {
          check();
          const save = () => {
            try {
              check();
              download(
                buffer,
                `${name()}.xlsx`,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              );
            } catch (e) {
              notice(String(e));
            }
          };
          const warning =
            buffer.byteLength > MAX_FOREIGN_FILE_BYTES
              ? 'This XLSX exceeds the 20 MiB import limit of this playground and cannot be reopened here.'
              : undefined;
          if (report.issues.length || warning)
            showReport(report, save, undefined, warning, file.sheets);
          else save();
        } catch (e) {
          notice(String(e));
        }
      };
      worker.postMessage(file);
    } catch (e) {
      stopExport();
      if (!disposed) notice(String(e));
    }
  });
  bindings.on($('import'), 'click', () => $<HTMLInputElement>('file').click());
  let nativeImport: AbortController | undefined;
  let importWorker: Worker | undefined;
  let importTimer: ReturnType<typeof setTimeout> | undefined;
  let generation = 0;
  function cancelImport() {
    generation++;
    nativeImport?.abort();
    nativeImport = undefined;
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
    cancelImport();
    if (!/\.(json|opensheet)$/i.test(file.name) && file.size > MAX_FOREIGN_FILE_BYTES) {
      notice('Files must be smaller than 20 MiB');
      return;
    }
    if (/\.(json|opensheet)$/i.test(file.name)) {
      const controller = new AbortController(),
        request = generation;
      nativeImport = controller;
      importButton.disabled = true;
      importButton.textContent = 'Reading…';
      const active = () => nativeImport === controller && request === generation && !disposed;
      try {
        if (/\.json$/i.test(file.name)) await app.loadJSON(file, { signal: controller.signal });
        else await app.loadBinary(file, { signal: controller.signal });
        if (!active()) return;
        state.dirty = true;
        state.importedWorkbook = true;
        $<HTMLInputElement>('document-name').value = file.name.replace(
          /\.(opensheet\.json|opensheet|json)$/i,
          '',
        );
        refreshPersistence(app, state);
        refresh();
        await app.getGrid()?.ready();
      } catch (error) {
        if (active()) notice(String(error));
      } finally {
        if (active()) {
          nativeImport = undefined;
          importButton.disabled = false;
          importButton.innerHTML = importLabel;
        }
      }
      return;
    }
    const request = generation;
    const originalBook = app.getWorkbook();
    const originalRevision = originalBook.revision;
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
    worker.onmessage = (
      event: MessageEvent<{
        error?: string;
        file: WorkbookFile;
        report?: CompatibilityReport;
      }>,
    ) => {
      if (!active()) return;
      stop();
      if (event.data.error) {
        notice(event.data.error);
        return;
      }
      const current = () =>
        request === generation &&
        !disposed &&
        app.getWorkbook() === originalBook &&
        originalBook.revision === originalRevision;
      if (!current()) {
        notice('Workbook changed during import. Import again.');
        return;
      }
      if ($<HTMLDialogElement>('dialog').open) {
        notice('Close the current dialog and import again.');
        return;
      }
      const report = event.data.report;
      const needsReview =
        !!report &&
        (report.summary.approximated > 0 ||
          report.summary.dropped > 0 ||
          report.summary.blocked > 0);
      const load = async () => {
        if (!current()) {
          notice('Workbook changed during import. Import again.');
          return;
        }
        const controller = new AbortController();
        nativeImport = controller;
        b.disabled = true;
        b.textContent = 'Reading…';
        const loading = () => nativeImport === controller && request === generation && !disposed;
        try {
          await app.load(event.data.file, { signal: controller.signal });
          if (!loading()) return;
          state.dirty = true;
          state.importedWorkbook = true;
          $<HTMLInputElement>('document-name').value = file.name.replace(
            /\.(xlsx|csv|tsv|json)$/i,
            '',
          );
          refreshPersistence(app, state);
          refresh();
          if (!needsReview && report?.issues.length) {
            if ($<HTMLDialogElement>('dialog').open)
              notice('Workbook imported; compatibility details could not be shown.');
            else showReport(report, undefined, undefined, undefined, event.data.file.sheets);
          }
        } catch (e) {
          if (loading()) notice(String(e));
        } finally {
          if (loading()) {
            nativeImport = undefined;
            b.disabled = false;
            b.innerHTML = importLabel;
          }
        }
      };
      if (needsReview)
        showReport(
          report,
          () => void load(),
          'Import converted workbook',
          undefined,
          event.data.file.sheets,
        );
      else void load();
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
      stopExport();
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

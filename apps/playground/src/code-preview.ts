import { streamExport, type DelimitedExportReport, type Format } from '@opensheetjs/formats';
import type { OpenSheet } from 'opensheet';
import { createBindings } from './bindings.js';
import { $ } from './dom.js';
import type { Feedback } from './feedback.js';
export function createCodePreview(
  app: OpenSheet,
  { notice, textModal }: Pick<Feedback, 'notice' | 'textModal'>,
  download: (data: BlobPart, name: string, type?: string) => void,
) {
  const bindings = createBindings();
  const name = () => $<HTMLInputElement>('document-name').value.trim() || 'workbook';
  let format: Format = 'latex',
    sourceText = '',
    refreshTimer: ReturnType<typeof setTimeout> | undefined;
  let generation = 0;
  function refresh() {
    const request = ++generation;
    clearTimeout(refreshTimer);
    $<HTMLButtonElement>('copy-code').disabled = true;
    $('format-preview').setAttribute('aria-busy', 'true');
    $('source').textContent = 'Updating preview…';
    refreshTimer = setTimeout(async () => {
      try {
        const s = app.selection.get()!;
        const config = {
          sheetId: s.sheetId,
          format,
          ...($<HTMLInputElement>('export-selection').checked ? { range: s } : {}),
        };
        const used = app.getWorkbook().usedRange(s.sheetId);
        const range = $<HTMLInputElement>('export-selection').checked
          ? s
          : { sheetId: s.sheetId, ...used };
        const previewRange = {
          ...range,
          endRow: Math.min(range.endRow, range.startRow + 200),
          endColumn: Math.min(range.endColumn, range.startColumn + 32),
        };
        let text = '';
        let report: DelimitedExportReport | undefined;
        for await (const chunk of streamExport(
          app.getWorkbook(),
          { ...config, range: previewRange },
          {
            onReport: (result) => {
              report = result;
            },
          },
        ))
          text += chunk;
        if (request !== generation) return;
        sourceText = text;
        $('source').textContent =
          sourceText.length > 30000
            ? sourceText.slice(0, 30000) +
              '\n… Preview truncated; download includes the full result.'
            : sourceText;
        const clipped =
          previewRange.endRow < range.endRow || previewRange.endColumn < range.endColumn;
        const limited = clipped || sourceText.length > 30000;
        const delimited = format === 'csv' || format === 'tsv';
        const affected = report?.sanitized
          ? ` Preview: ${report.sanitized} ${report.sanitized === 1 ? 'cell' : 'cells'} prefixed (${report.sample.join(', ')}${report.truncated ? ', …' : ''}).`
          : '';
        $('export-note').textContent = delimited
          ? `Text with formula or control prefixes is prefixed with an apostrophe; spreadsheet imports may change values.${affected}${limited ? ' Preview is limited; download includes the full result.' : ''}`
          : limited
            ? 'Preview is limited; download includes the full result.'
            : 'Ready to paste into your next project.';
        $('format-preview').removeAttribute('aria-busy');
        $<HTMLButtonElement>('copy-code').disabled = false;
      } catch (e) {
        if (request !== generation) return;
        $('format-preview').removeAttribute('aria-busy');
        $('source').textContent = e instanceof Error ? e.message : String(e);
        $('export-note').textContent = 'Preview unavailable.';
        sourceText = '';
      }
    }, 120);
  }
  const formatTabs = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-format]'));
  function selectFormat(tab: HTMLButtonElement) {
    format = tab.dataset.format as Format;
    for (const item of formatTabs) {
      const selected = item === tab;
      item.classList.toggle('selected', selected);
      item.setAttribute('aria-selected', String(selected));
      item.tabIndex = selected ? 0 : -1;
    }
    $('format-preview').setAttribute('aria-labelledby', tab.id);
    refresh();
  }
  for (const [index, tab] of formatTabs.entries()) {
    bindings.on(tab, 'click', () => selectFormat(tab));
    bindings.on(tab, 'keydown', (event) => {
      let next: number;
      if (event.key === 'ArrowRight') next = (index + 1) % formatTabs.length;
      else if (event.key === 'ArrowLeft')
        next = (index + formatTabs.length - 1) % formatTabs.length;
      else if (event.key === 'Home') next = 0;
      else if (event.key === 'End') next = formatTabs.length - 1;
      else return;
      event.preventDefault();
      selectFormat(formatTabs[next]);
      formatTabs[next].focus();
    });
  }
  bindings.on($<HTMLInputElement>('export-selection'), 'change', refresh);
  bindings.on($('download-code'), 'click', async () => {
    try {
      const book = app.getWorkbook(),
        selection = app.selection.get()!,
        exportFormat = format,
        exportName = name();
      const parts: string[] = [];
      let report: DelimitedExportReport | undefined;
      for await (const chunk of streamExport(
        book,
        {
          sheetId: selection.sheetId,
          format: exportFormat,
          ...($<HTMLInputElement>('export-selection').checked ? { range: selection } : {}),
        },
        {
          onReport: (result) => {
            report = result;
          },
        },
      ))
        parts.push(chunk);
      download(
        new Blob(parts),
        `${exportName}.${exportFormat === 'latex' ? 'tex' : exportFormat === 'markdown' ? 'md' : exportFormat}`,
      );
      if (report) {
        const examples = report.sample.slice(0, 5);
        const addresses = examples.length
          ? ` (${examples.join(', ')}${report.sanitized > examples.length ? ', …' : ''})`
          : '';
        notice(
          `Downloaded ${report.format.toUpperCase()}: ${report.sanitized} ${report.sanitized === 1 ? 'cell' : 'cells'} prefixed${addresses}.`,
        );
      }
    } catch (error) {
      notice(String(error));
    }
  });
  bindings.on($('copy-code'), 'click', async () => {
    try {
      await navigator.clipboard.writeText(sourceText);
      notice('Preview copied to clipboard');
    } catch {
      textModal('Copy this code', sourceText);
    }
  });
  return {
    refresh,
    dispose() {
      ++generation;
      clearTimeout(refreshTimer);
      bindings.dispose();
    },
  };
}

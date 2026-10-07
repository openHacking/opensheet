import { streamExport, exportRange, type Format } from '@opensheetjs/formats';
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
        const result = exportRange(await app.getWorkbook().toJSON(previewRange), {
          ...config,
          range: previewRange,
        });
        if (request !== generation) return;
        sourceText = result.text;
        $('source').textContent =
          sourceText.length > 30000
            ? sourceText.slice(0, 30000) +
              '\n… Preview truncated; download includes the full result.'
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
  for (const b of document.querySelectorAll<HTMLButtonElement>('[data-format]'))
    bindings.on(b, 'click', () => {
      format = b.dataset.format as Format;
      for (const x of document.querySelectorAll('[data-format]')) {
        x.classList.toggle('selected', x === b);
        x.setAttribute('aria-selected', String(x === b));
      }
      refresh();
    });
  bindings.on($<HTMLInputElement>('export-selection'), 'change', refresh);
  bindings.on($('download-code'), 'click', async () => {
    try {
      const book = app.getWorkbook(),
        selection = app.selection.get()!;
      const parts: string[] = [];
      for await (const chunk of streamExport(book, {
        sheetId: selection.sheetId,
        format,
        ...($<HTMLInputElement>('export-selection').checked ? { range: selection } : {}),
      }))
        parts.push(chunk);
      download(
        new Blob(parts),
        `${name()}.${format === 'latex' ? 'tex' : format === 'markdown' ? 'md' : format}`,
      );
    } catch (error) {
      notice(String(error));
    }
  });
  bindings.on($('copy-code'), 'click', async () => {
    try {
      await navigator.clipboard.writeText(sourceText);
      notice('Code copied to clipboard');
    } catch {
      textModal('Copy this code', sourceText);
    }
  });
  return {
    refresh,
    dispose() {
      clearTimeout(refreshTimer);
      bindings.dispose();
    },
  };
}

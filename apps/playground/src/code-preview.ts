import { exportRange, type Format } from '@opensheetjs/formats';
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
  function refresh() {
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
  bindings.on($('download-code'), 'click', () =>
    download(
      sourceText,
      `${name()}.${format === 'latex' ? 'tex' : format === 'markdown' ? 'md' : format}`,
    ),
  );
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

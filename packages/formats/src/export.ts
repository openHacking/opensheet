import { OpenSheetError, FileReader, type WorkbookFile } from '@opensheetjs/core';
import { renderDelimited } from './generators/delimited.js';
import { renderHtml } from './generators/html.js';
import { renderLatex } from './generators/latex.js';
import { renderMarkdown } from './generators/markdown.js';
import { prepareRange } from './prepare.js';
import type { ExportOptions, ExportResult } from './types.js';
export function exportRange(file: WorkbookFile, config: ExportOptions): ExportResult {
  if (!['csv', 'tsv', 'markdown', 'html', 'latex'].includes(config.format))
    throw new OpenSheetError('INVALID_ARGUMENT', 'Unknown export format');
  const book = new FileReader(file);
  try {
    const prepared = prepareRange(book, config);
    switch (config.format) {
      case 'csv':
      case 'tsv':
        return renderDelimited(prepared, config);
      case 'markdown':
        return renderMarkdown(prepared);
      case 'html':
        return renderHtml(prepared);
      case 'latex':
        return renderLatex(prepared);
    }
  } finally {
    book.dispose();
  }
}

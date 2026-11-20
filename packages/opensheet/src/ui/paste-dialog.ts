import { address, OpenSheetError, type Selection, type Workbook } from '@opensheetjs/core';
import { pasteClipboard, type PasteOptions } from '@opensheetjs/renderer';
import { currentWorkbookDialog, type WorkbookDialogFeedback } from './workbook-dialog.js';

export function showPasteDialog(
  book: Workbook,
  selection: Selection,
  { modal, notice, isCurrent, isTargetCurrent = isCurrent }: WorkbookDialogFeedback,
  select: (selection: Selection) => void,
  initialMode: NonNullable<PasteOptions['mode']> = 'all',
) {
  const target = { ...selection },
    baseRevision = book.revision;
  const form = document.createElement('form');
  form.className = 'sort-form annotation-form';
  const info = document.createElement('p');
  info.textContent = `Paste into the box below, choose what to keep, then apply at ${address(target.startRow, target.startColumn)}. Nothing is written until Apply.`;
  const mode = document.createElement('select');
  mode.tabIndex = 0;
  mode.setAttribute('aria-label', 'Paste mode');
  for (const [value, label] of [
    ['all', 'Values, formulas and available formats'],
    ['values', 'Values only'],
    ['formulas', 'Formula inputs only'],
    ['formats', 'Formats only'],
  ])
    mode.add(new Option(label, value));
  const description = document.createElement('p');
  const descriptions: Record<string, string> = {
    all: 'HTML tables include supported formats. Plain text preserves destination formats.',
    values:
      'Keeps destination formats. OpenSheet copies use captured values; external text beginning with = stays literal.',
    formulas:
      'Keeps destination formats. OpenSheet copies relocate relative references; external formula text uses the references as pasted.',
    formats:
      'Requires an HTML table. Replaces supported formats while preserving destination contents, notes and links.',
  };
  mode.onchange = () => {
    description.textContent = descriptions[mode.value];
  };
  mode.value = initialMode;
  mode.onchange(new Event('change'));
  const box = document.createElement('textarea');
  box.setAttribute('aria-label', 'Paste clipboard here');
  box.rows = 5;
  box.readOnly = true;
  box.placeholder = 'Focus here and press Ctrl+V or Command+V';
  const status = document.createElement('p');
  status.setAttribute('role', 'alert');
  const apply = document.createElement('button');
  apply.tabIndex = 0;
  apply.type = 'submit';
  apply.textContent = 'Apply paste';
  apply.disabled = true;
  let data: { text: string; html?: string } | undefined;
  box.onpaste = (event) => {
    event.preventDefault();
    if (!currentWorkbookDialog(form, status, isTargetCurrent)) return;
    if (book.isReadOnly || !event.clipboardData) return;
    if (!event.clipboardData.types.some((type) => type === 'text/plain' || type === 'text/html')) {
      data = undefined;
      box.value = '';
      apply.disabled = true;
      status.textContent = 'Clipboard has no supported text or HTML table.';
      return;
    }
    const text = event.clipboardData.getData('text/plain'),
      html = event.clipboardData.getData('text/html');
    if (text.length > 2 * 1024 * 1024 || html.length > 2 * 1024 * 1024) {
      data = undefined;
      apply.disabled = true;
      status.textContent = 'Paste is too large';
      return;
    }
    data = { text, html };
    box.value = text.slice(0, 4000) || (html ? 'HTML table captured' : '(Empty cell)');
    status.textContent =
      text.length > 4000
        ? 'Clipboard captured. Preview shows the first 4,000 characters.'
        : 'Clipboard captured. Choose a mode and apply.';
    apply.disabled = false;
  };
  form.onsubmit = async (event) => {
    event.preventDefault();
    if (!currentWorkbookDialog(form, status, isTargetCurrent)) return;
    if (!data) return;
    apply.disabled = mode.disabled = box.disabled = true;
    status.textContent = 'Pasting…';
    try {
      const result = await pasteClipboard(book, target, data, {
        mode: mode.value as PasteOptions['mode'],
        baseRevision,
      });
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      select(result);
      form.closest('dialog')?.close();
      notice('Clipboard pasted');
    } catch (reason) {
      if (!currentWorkbookDialog(form, status, isCurrent)) return;
      if (reason instanceof OpenSheetError && reason.code === 'REVISION_CONFLICT') {
        status.textContent = 'Workbook changed. Close this dialog and select the target again.';
        return;
      }
      status.textContent = reason instanceof Error ? reason.message : String(reason);
      try {
        apply.disabled = mode.disabled = box.disabled = book.isReadOnly;
      } catch {
        /* Disposed workbook stays disabled. */
      }
    }
  };
  if (book.isReadOnly) {
    box.disabled = mode.disabled = true;
    status.textContent = 'Read only. Pasting is disabled.';
  }
  form.append(info, mode, description, box, status, apply);
  modal('Paste special', form);
  if (!book.isReadOnly) box.focus();
}

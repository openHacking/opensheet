import { OpenSheetError } from '@opensheetjs/core';
export async function readClipboard() {
  const clipboard = navigator.clipboard;
  if (!clipboard?.read) return;
  const types = ['text/html', 'text/plain'];
  let item: ClipboardItem | undefined;
  let html: Blob | undefined;
  let text: Blob | undefined;
  try {
    const items = await clipboard.read();
    item =
      items.find((entry) => entry.types.includes(types[0])) ??
      items.find((entry) => entry.types.includes(types[1]));
    if (!item)
      throw new OpenSheetError(
        'INVALID_ARGUMENT',
        'Clipboard has no supported text or HTML table.',
      );
    if (item.types.includes('text/html')) html = await item.getType('text/html');
    if (item.types.includes('text/plain')) text = await item.getType('text/plain');
  } catch (error) {
    if (error instanceof OpenSheetError) throw error;
    return;
  }
  if ((html?.size ?? 0) > 2 * 1024 * 1024 || (text?.size ?? 0) > 2 * 1024 * 1024)
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste is too large');
  return { html: await html?.text(), text: (await text?.text()) ?? '' };
}

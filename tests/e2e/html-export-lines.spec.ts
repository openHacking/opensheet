import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('HTML export preserves line breaks and treats cell markup as text', async ({ page }) => {
  const formatsModule = `/@fs${fileURLToPath(new URL('../../packages/formats/src/index.ts', import.meta.url))}`;
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (modulePath) => {
    const book = window.opensheet.getWorkbook();
    const sheet = await book.addSheet('HTML lines', { rows: 1, columns: 1 });
    await sheet.range('A1').setValues([['first\rsecond\r\n<img src=x>\nlast & end']]);
    const { streamExport } = await import(/* @vite-ignore */ modulePath);
    let html = '';
    for await (const chunk of streamExport(book, { sheetId: sheet.id, format: 'html' }))
      html += chunk;
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    const cell = parsed.querySelector('th')!;
    return {
      breaks: cell.querySelectorAll('br').length,
      images: parsed.querySelectorAll('img').length,
      text: cell.textContent,
    };
  }, formatsModule);
  expect(result).toEqual({ breaks: 3, images: 0, text: 'firstsecond<img src=x>last & end' });
});

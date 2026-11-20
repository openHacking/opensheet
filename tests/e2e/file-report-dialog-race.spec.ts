import { workbookTool } from './workbook-tools.js';
import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';

for (const operation of ['import', 'export'] as const) {
  test(`late ${operation} report does not replace an open dialog`, async ({ page }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(async (operation) => {
      const book = window.opensheet.getWorkbook();
      if (operation === 'export')
        await book
          .getSheets()[0]
          .range('A1')
          .setValues([['']]);
      const NativeWorker = window.Worker;
      window.Worker = new Proxy(NativeWorker, {
        construct(target, args) {
          const worker = Reflect.construct(target, args) as Worker;
          const filename = String(args[0]);
          if (filename.includes(operation === 'import' ? 'file.worker' : 'xlsx-export.worker')) {
            worker.addEventListener('message', (event) => {
              event.stopImmediatePropagation();
              (window as any).releaseFileReport = () => worker.onmessage?.call(worker, event);
            });
          }
          return worker;
        },
      });
      return { id: book.id, revision: book.revision };
    }, operation);
    const downloads: string[] = [];
    page.on('download', (download) => downloads.push(download.suggestedFilename()));
    if (operation === 'export') await page.locator('#export-xlsx').click();
    else {
      const workbook = XLSX.utils.book_new(),
        sheet = XLSX.utils.aoa_to_sheet([[1]]);
      sheet.A1.f = 'UNKNOWNFUNC(1)';
      XLSX.utils.book_append_sheet(workbook, sheet, 'Incoming');
      await page.locator('#file').setInputFiles({
        name: 'incoming.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        buffer: XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }),
      });
    }
    await expect
      .poll(() => page.evaluate(() => typeof (window as any).releaseFileReport))
      .toBe('function');
    await workbookTool(page, 'sort-custom');
    const dialog = page.getByRole('dialog');
    await expect(dialog.getByRole('button', { name: 'Apply sort' })).toBeVisible();
    await page.evaluate(() => (window as any).releaseFileReport());
    await expect(dialog.getByRole('button', { name: 'Apply sort' })).toBeVisible();
    await expect(page.locator('#notice')).toContainText('Close the current dialog');
    expect(downloads).toEqual([]);
    expect(
      await page.evaluate(() => ({
        id: window.opensheet.getWorkbook().id,
        revision: window.opensheet.getWorkbook().revision,
      })),
    ).toEqual(before);
  });
}

test('a replaced report cannot trigger its detached download button', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['']]);
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog.getByRole('button', { name: 'Download converted XLSX' })).toBeVisible();
  await page.evaluate(() => {
    const button = document.querySelector<HTMLButtonElement>('#dialog-content button')!;
    const createObjectURL = URL.createObjectURL.bind(URL);
    (window as any).downloadUrls = 0;
    URL.createObjectURL = (value) => {
      (window as any).downloadUrls++;
      return createObjectURL(value);
    };
    document
      .querySelector('#dialog-content')!
      .replaceChildren(document.createTextNode('Replacement'));
    button.click();
  });
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('Replacement');
  expect(downloads).toEqual([]);
  expect(await page.evaluate(() => (window as any).downloadUrls)).toBe(0);
});

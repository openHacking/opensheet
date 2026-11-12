import { expect, test } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(() => {
    const NativeWorker = window.Worker;
    window.Worker = new Proxy(NativeWorker, {
      construct(target, args) {
        const worker = Reflect.construct(target, args) as Worker;
        if (String(args[0]).includes('xlsx-export.worker')) {
          (window as any).exportTestWorker = worker;
          const terminate = worker.terminate.bind(worker);
          worker.terminate = () => {
            (window as any).exportWorkerTerminated = true;
            terminate();
          };
          worker.addEventListener('message', (event) => {
            event.stopImmediatePropagation();
            (window as any).releaseExport = () => worker.onmessage?.call(worker, event);
            (window as any).releaseLargeExport = () =>
              worker.onmessage?.call(
                worker,
                new MessageEvent('message', {
                  data: { ...event.data, buffer: new ArrayBuffer(20 * 1024 * 1024 + 1) },
                }),
              );
          });
        }
        return worker;
      },
    });
  });
});

test('warns before downloading an XLSX larger than the playground import limit', async ({
  page,
}) => {
  await page.locator('#export-xlsx').click();
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseLargeExport))
    .toBe('function');
  await page.evaluate(() => (window as any).releaseLargeExport());
  await expect(page.getByRole('dialog', { name: 'Compatibility report' })).toContainText(
    'exceeds the 20 MiB import limit of this playground and cannot be reopened here',
  );
});

test('rejects a completed export when the workbook changed during worker processing', async ({
  page,
}) => {
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseExport))
    .toBe('function');
  await expect(page.locator('#export-xlsx')).toBeDisabled();
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['new value']]);
    (window as any).releaseExport();
  });
  await expect(page.locator('#notice')).toContainText('Workbook changed; export again');
  await expect(page.locator('#export-xlsx')).toBeEnabled();
  expect(downloads).toEqual([]);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['new value']]);
});

test('rejects a reviewed export when the workbook changes before download', async ({ page }) => {
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['']]);
  });
  await page.locator('#export-xlsx').click();
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releaseExport))
    .toBe('function');
  await page.evaluate(() => (window as any).releaseExport());
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Empty text may be lost');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['new value']]);
  });
  await dialog.getByRole('button', { name: 'Download converted XLSX' }).click();
  await expect(page.locator('#notice')).toContainText('Workbook changed; export again');
  expect(downloads).toEqual([]);
});

for (const [type, message] of [
  ['error', 'Could not prepare XLSX'],
  ['messageerror', 'Could not receive XLSX'],
] as const) {
  test(`recovers the export control after a worker ${type}`, async ({ page }) => {
    await page.locator('#export-xlsx').click();
    await expect
      .poll(() => page.evaluate(() => typeof (window as any).exportTestWorker?.onmessage))
      .toBe('function');
    await page.evaluate((type) => {
      const worker = (window as any).exportTestWorker as Worker;
      if (type === 'messageerror') worker.dispatchEvent(new MessageEvent(type));
      else worker.dispatchEvent(new ErrorEvent(type));
    }, type);
    await expect(page.locator('#notice')).toContainText(message);
    await expect(page.locator('#export-xlsx')).toBeEnabled();
    expect(await page.evaluate(() => (window as any).exportWorkerTerminated)).toBe(true);
  });
}

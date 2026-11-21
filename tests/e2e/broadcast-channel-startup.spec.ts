import { expect, test } from '@playwright/test';

test('denied tab synchronization releases the writer and allows the saved workbook to reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (modulePath) => {
    const { openWorkbook } = await import(/* @vite-ignore */ modulePath);
    const original = window.opensheet.getWorkbook();
    await original
      .getSheets()[0]
      .range('A1')
      .setValues([[12345]]);
    const id = original.id;
    const revision = original.revision;
    await original.close();

    const Channel = globalThis.BroadcastChannel;
    const failures: { code?: string; name: string; message: string }[] = [];
    try {
      for (const error of [
        new DOMException('Blocked', 'SecurityError'),
        new DOMException('Blocked', 'NotAllowedError'),
        new TypeError('Unexpected constructor defect'),
      ]) {
        globalThis.BroadcastChannel = class extends Channel {
          constructor(name: string) {
            super(name);
            this.close();
            throw error;
          }
        };
        try {
          await openWorkbook(id);
          throw new Error('Expected channel construction to fail');
        } catch (failure) {
          const observed = failure as Error & { code?: string };
          failures.push({ code: observed.code, name: observed.name, message: observed.message });
        }
      }
    } finally {
      globalThis.BroadcastChannel = Channel;
    }
    const reopened = await openWorkbook(id);
    const recovered = {
      revision: reopened.revision,
      writer: !reopened.isReadOnly,
      value: await reopened.getSheets()[0].range('A1').getValues(),
    };
    await reopened.close();
    return { failures, revision, recovered };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);

  expect(result.failures[0]).toMatchObject({
    code: 'BROADCAST_CHANNEL_UNAVAILABLE',
    message: expect.stringContaining('reopen the workbook'),
  });
  expect(result.failures[1]).toMatchObject({ code: 'BROADCAST_CHANNEL_UNAVAILABLE' });
  expect(result.failures[2]).toMatchObject({
    name: 'TypeError',
    message: 'Unexpected constructor defect',
  });
  expect(result.recovered).toEqual({ revision: result.revision, writer: true, value: [[12345]] });
});

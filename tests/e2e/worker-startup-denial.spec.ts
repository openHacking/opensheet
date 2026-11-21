import { expect, test } from '@playwright/test';

test('denied Worker construction reports a remedy and permits reopening saved data', async ({
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
      .setValues([[54321]]);
    const { id, revision } = original;
    await original.close();
    const WorkerConstructor = globalThis.Worker;
    let failure: { code?: string; message?: string } | undefined;
    try {
      globalThis.Worker = class {
        constructor() {
          throw new DOMException('Blocked by test policy', 'SecurityError');
        }
      } as unknown as typeof Worker;
      try {
        await openWorkbook(id);
      } catch (error) {
        const observed = error as Error & { code?: string };
        failure = { code: observed.code, message: observed.message };
      }
    } finally {
      globalThis.Worker = WorkerConstructor;
    }
    const reopened = await openWorkbook(id);
    const recovered = {
      revision: reopened.revision,
      writer: !reopened.isReadOnly,
      value: await reopened.getSheets()[0].range('A1').getValues(),
    };
    await reopened.close();
    return { failure, revision, recovered };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.failure).toMatchObject({
    code: 'WORKER_UNAVAILABLE',
    message: expect.stringContaining('security policy'),
  });
  expect(result.recovered).toEqual({ revision: result.revision, writer: true, value: [[54321]] });
});

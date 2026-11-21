import { expect, test } from '@playwright/test';

test('reports initialization failures in an opaque sandbox without modifying a saved workbook', async ({
  page,
}) => {
  await page.route('**/*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      headers: { ...response.headers(), 'access-control-allow-origin': '*' },
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (modulePath) => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[6842]]);
    const revision = book.revision;
    const source = `<!doctype html><script type="module">
      let workerStarts = 0;
      const NativeWorker = globalThis.Worker;
      globalThis.Worker = class extends NativeWorker {
        constructor(...args) { workerStarts++; super(...args); }
      };
      try {
        const { createWorkbook } = await import(${JSON.stringify(new URL(modulePath, location.href).href)});
        const workbook = await createWorkbook({ sheets: [{ name:'Sandbox', rows:1, columns:1 }] });
        await workbook.deleteStorage();
        parent.postMessage({ sandboxProbe:true, success:true, workerStarts }, '*');
      } catch (error) {
        parent.postMessage({ sandboxProbe:true, workerStarts, code:error?.code, name:error?.name, message:String(error?.message ?? error) }, '*');
      }
    <\/script>`;
    const probe = async (sameOrigin: boolean) => {
      const frame = document.createElement('iframe');
      frame.sandbox.add('allow-scripts');
      if (sameOrigin) frame.sandbox.add('allow-same-origin');
      frame.srcdoc = source;
      const reply = new Promise<any>((resolve, reject) => {
        const timeout = setTimeout(() => {
          window.removeEventListener('message', listener);
          reject(new Error('Sandbox did not respond'));
        }, 15000);
        const listener = (event: MessageEvent) => {
          if (event.source !== frame.contentWindow || !event.data?.sandboxProbe) return;
          clearTimeout(timeout);
          window.removeEventListener('message', listener);
          resolve(event.data);
        };
        window.addEventListener('message', listener);
      });
      document.body.append(frame);
      try {
        return await reply;
      } finally {
        frame.remove();
      }
    };
    const sandbox = await probe(false);
    const sameOrigin = await probe(true);
    const { openWorkbook } = await import(/* @vite-ignore */ modulePath);
    const reopened = await openWorkbook(book.id);
    try {
      return {
        sandbox,
        sameOrigin,
        revision,
        after: book.revision,
        value: await book.getSheets()[0].range('A1').getValues(),
        reopenedRevision: reopened.revision,
        reopenedValue: await reopened.getSheets()[0].range('A1').getValues(),
      };
    } finally {
      await reopened.close();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.sandbox).toMatchObject({
    code: 'STORAGE_UNAVAILABLE',
    workerStarts: 0,
    message: expect.stringContaining('opaque origin'),
  });
  expect(result.sameOrigin).toMatchObject({ success: true, workerStarts: 1 });
  expect(result.after).toBe(result.revision);
  expect(result.value).toEqual([[6842]]);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.reopenedValue).toEqual([[6842]]);
});

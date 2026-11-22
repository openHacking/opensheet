import { expect, test } from '@playwright/test';

for (const phase of ['active', 'queued'] as const) {
  test(`cancels only the ${phase} cell read and preserves later work`, async ({ page }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let holdNextCopyRead = false, releaseCopyRead;
addEventListener('message', ({data}) => {
  if (data.armCopyRead) { holdNextCopyRead = true; postMessage({copyReadArmed: true}); }
  if (data.method === 'cancelRequest') postMessage({copyCancelObserved: data.args[0]});
  if (data.releaseCopyRead && releaseCopyRead) { const release = releaseCopyRead; releaseCopyRead = undefined; release(); }
});
const copyOriginalGet = IDBObjectStore.prototype.get;
IDBObjectStore.prototype.get = function(...args) {
  const req = copyOriginalGet.apply(this, args);
  if (!holdNextCopyRead || this.name !== 'records') return req;
  holdNextCopyRead = false;
  return {
    get result() { return req.result; }, get error() { return req.error; },
    set onsuccess(fn) { req.onsuccess = (event) => { releaseCopyRead = () => fn(event); postMessage({copyReadStalled: true}); }; },
    set onerror(fn) { req.onerror = fn; }
  };
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(
      async ({ coreUrl, phase }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ coreUrl);
        let book = await createWorkbook(
          { sheets: [{ name: 'Data', rows: 2, columns: 1 }] },
          { cacheBytes: 0 },
        );
        try {
          let sheet = book.getSheets()[0];
          await sheet.range('A1').setValues([[7]]);
          const worker = (book as any).worker as Worker;
          const wait = (field: string) =>
            new Promise<any>((resolve, reject) => {
              const timer = setTimeout(
                () => reject(new Error(`Worker did not report ${field}`)),
                10000,
              );
              worker.addEventListener('message', function receive({ data }) {
                if (data[field] !== undefined) {
                  clearTimeout(timer);
                  worker.removeEventListener('message', receive);
                  resolve(data[field]);
                }
              });
            });
          const armed = wait('copyReadArmed');
          worker.postMessage({ method: 'cancelRequest', args: [-1], armCopyRead: true });
          await armed;
          const stalled = wait('copyReadStalled');
          const firstController = new AbortController();
          const first = book
            .readCells(
              sheet.id,
              { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
              { signal: firstController.signal },
            )
            .then(
              (value: any) => value.cells[0].value,
              (error: { code: string }) => error.code,
            );
          await stalled;
          let second: Promise<unknown> | undefined;
          const cancelled = wait('copyCancelObserved');
          if (phase === 'active') firstController.abort();
          else {
            const controller = new AbortController();
            second = book
              .readCells(
                sheet.id,
                { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
                { signal: controller.signal },
              )
              .then(
                () => 'completed',
                (error: { code: string }) => error.code,
              );
            controller.abort();
          }
          await cancelled;
          const write = sheet.range('A1').setValues([[9]]);
          worker.postMessage({ method: 'cancelRequest', args: [-1], releaseCopyRead: true });
          const firstResult = await first;
          const secondResult = second ? await second : null;
          await write;
          const current = await sheet.range('A1').getValues();
          const revision = book.revision;
          const id = book.id;
          await book.close();
          book = await openWorkbook(id, { cacheBytes: 0 });
          sheet = book.getSheets()[0];
          return {
            firstResult,
            secondResult,
            current,
            revision,
            reopenedRevision: book.revision,
            reopened: await sheet.range('A1').getValues(),
          };
        } finally {
          await book.deleteStorage();
        }
      },
      { coreUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`, phase },
    );
    expect(result.firstResult).toEqual(phase === 'active' ? 'ABORTED' : 7);
    expect(result.secondResult).toEqual(phase === 'queued' ? 'ABORTED' : null);
    expect(result.current).toEqual([[9]]);
    expect(result.reopened).toEqual([[9]]);
    expect(result.reopenedRevision).toBe(result.revision);
  });
}

test('clipboard denial cancels its active Worker read and leaves later edits usable', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let holdClipboardRead = false, resumeClipboardRead;
addEventListener('message', ({data}) => {
  if (data.armClipboardRead) { holdClipboardRead = true; postMessage({clipboardReadArmed: true}); }
  if (data.method === 'cancelRequest' && data.args[0] > 0) postMessage({clipboardCancelObserved: data.args[0]});
  if (data.releaseClipboardRead && resumeClipboardRead) { const resume = resumeClipboardRead; resumeClipboardRead = undefined; resume(); }
});
const clipboardOriginalGet = IDBObjectStore.prototype.get;
IDBObjectStore.prototype.get = function(...args) {
  const req = clipboardOriginalGet.apply(this, args);
  if (!holdClipboardRead || this.name !== 'records') return req;
  holdClipboardRead = false;
  return {
    get result() { return req.result; }, get error() { return req.error; },
    set onsuccess(fn) { req.onsuccess = (event) => { resumeClipboardRead = () => fn(event); postMessage({clipboardReadStalled: true}); }; },
    set onerror(fn) { req.onerror = fn; }
  };
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async ({ coreUrl, clipboardUrl }) => {
      const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ coreUrl);
      const { writeClipboard } = await import(/* @vite-ignore */ clipboardUrl);
      let book = await createWorkbook(
        { sheets: [{ name: 'Data', rows: 2, columns: 1 }] },
        { cacheBytes: 0 },
      );
      const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
      const item = Object.getOwnPropertyDescriptor(window, 'ClipboardItem');
      try {
        let sheet = book.getSheets()[0];
        await sheet.range('A1').setValues([[7]]);
        const worker = (book as any).worker as Worker;
        const wait = (field: string) =>
          new Promise<any>((resolve, reject) => {
            const timer = setTimeout(
              () => reject(new Error(`Worker did not report ${field}`)),
              10000,
            );
            worker.addEventListener('message', function receive({ data }) {
              if (data[field] !== undefined) {
                clearTimeout(timer);
                worker.removeEventListener('message', receive);
                resolve(data[field]);
              }
            });
          });
        const armed = wait('clipboardReadArmed');
        worker.postMessage({ method: 'cancelRequest', args: [-1], armClipboardRead: true });
        await armed;
        const stalled = wait('clipboardReadStalled');
        const cancellation = wait('clipboardCancelObserved');
        const denied = new DOMException('Clipboard denied', 'NotAllowedError');
        Object.defineProperty(navigator, 'clipboard', {
          configurable: true,
          value: {
            write: async () => {
              await stalled;
              throw denied;
            },
          },
        });
        Object.defineProperty(window, 'ClipboardItem', {
          configurable: true,
          value: class {
            constructor(_data: unknown) {}
          },
        });
        const operation = writeClipboard(sheet.range('A1')).then(
          () => 'success',
          (error: unknown) => (error === denied ? 'denied' : String(error)),
        );
        const beforeRelease = await operation;
        const requestId = await cancellation;
        const completed = new Promise<string>((resolve, reject) => {
          const timer = setTimeout(() => reject(new Error('Cancelled read did not settle')), 10000);
          worker.addEventListener('message', function receive({ data }) {
            if (data.id === requestId) {
              clearTimeout(timer);
              worker.removeEventListener('message', receive);
              resolve(data.error?.code ?? 'completed');
            }
          });
        });
        const edit = sheet.range('A1').setValues([[9]]);
        worker.postMessage({ method: 'cancelRequest', args: [-1], releaseClipboardRead: true });
        const readResult = await completed;
        await edit;
        const revision = book.revision;
        const id = book.id;
        await book.close();
        book = await openWorkbook(id, { cacheBytes: 0 });
        sheet = book.getSheets()[0];
        return {
          beforeRelease,
          readResult,
          revision,
          reopenedRevision: book.revision,
          value: await sheet.range('A1').getValues(),
        };
      } finally {
        if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
        else delete (navigator as any).clipboard;
        if (item) Object.defineProperty(window, 'ClipboardItem', item);
        else delete (window as any).ClipboardItem;
        await book.deleteStorage();
      }
    },
    {
      coreUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
      clipboardUrl: `/@fs${process.cwd()}/packages/renderer/src/copy-clipboard.ts`,
    },
  );
  expect(result.beforeRelease).toBe('denied');
  expect(result.readResult).toBe('ABORTED');
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.value).toEqual([[9]]);
});

test('pre-aborted cell reads do not dispatch to the Worker', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const worker = (book as any).worker as Worker;
    const original = worker.postMessage.bind(worker);
    let reads = 0;
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'readCells') reads++;
      return (original as any)(message, ...rest);
    }) as typeof worker.postMessage;
    const controller = new AbortController();
    controller.abort();
    try {
      const code = await book
        .readCells(
          sheet.id,
          { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
          { signal: controller.signal },
        )
        .then(
          () => 'completed',
          (error: { code: string }) => error.code,
        );
      return { code, reads };
    } finally {
      worker.postMessage = original;
    }
  });
  expect(result).toEqual({ code: 'ABORTED', reads: 0 });
});

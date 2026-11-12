import { expect, test } from '@playwright/test';

for (const surface of ['cell', 'formula'] as const) {
  test(`${surface} draft is retained when a pending insertion changes its cell identity`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    await page.evaluate(async (url) => {
      const { createWorkbookFile } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
      const sheet = app.getWorkbook().getSheets()[0];
      await sheet.range('A1').setValues([[42]]);
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
    }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
    if (surface === 'cell') {
      const grid = page.getByRole('grid', { exact: true });
      await grid.focus();
      await grid.press('Enter');
      await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('draft');
    } else {
      await page.getByLabel('Formula bar', { exact: true }).fill('draft');
    }
    await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let firstId = -1;
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute' && firstId === -1) firstId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === firstId) (window as any).editHeld = event;
        else handler.call(worker, event);
      };
      (window as any).editRelease = () => {
        handler.call(worker, (window as any).editHeld);
        worker.postMessage = post;
        worker.onmessage = handler;
      };
      (window as any).editInsert = book.getSheets()[0].insertRows(0);
    });
    await page.waitForFunction(() => !!(window as any).editHeld);
    if (surface === 'cell')
      await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Enter');
    else await page.getByLabel('Formula bar', { exact: true }).press('Enter');
    const recovery = page.getByRole('region', {
      name: surface === 'cell' ? 'Unconfirmed cell drafts' : 'Unconfirmed formula bar drafts',
    });
    await expect(recovery.locator('textarea')).toHaveValue('draft');
    await page.evaluate(async () => {
      (window as any).editRelease();
      await (window as any).editInsert;
    });
    const values = await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
    );
    expect(values).toEqual([[null], [42]]);
  });
}

test('a draft started without a loaded value falls back to a revision precondition', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const revision = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[42, 2]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    await app.getGrid()!.ready();
    (book as any).cache.clear();
    const element = app.getGrid()!.element;
    element.focus();
    element.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true }));
    await sheet.range('B1').setValues([[9]]);
    return book.revision;
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await expect(editor).toHaveValue('x');
  await editor.press('Enter');
  await expect(
    page.getByRole('region', { name: 'Unconfirmed cell drafts' }).locator('textarea'),
  ).toHaveValue('x');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
    ),
  ).toEqual([[42, 9]]);
});

for (const releaseReadFirst of [false, true]) {
  test(`formula draft handles an unloaded original value when read ${releaseReadFirst ? 'precedes' : 'follows'} a concurrent write`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    await page.evaluate(async (url) => {
      const { createWorkbookFile } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B1').setValues([[42, 2]]);
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
      await app.getGrid()!.ready();
      (book as any).cache.clear();
      const execute = book.execute.bind(book);
      (book as any).execute = (command: any) => {
        if (command.payload?.cells?.[0]?.input?.value === 'draft')
          (window as any).draftSubmission = structuredClone(command);
        return execute(command);
      };
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let readId = -1;
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'read' && message.args[2] === false && readId === -1)
          readId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === readId) (window as any).draftReadHeld = event;
        else handler.call(worker, event);
      };
      (window as any).releaseDraftRead = () => {
        handler.call(worker, (window as any).draftReadHeld);
        worker.postMessage = post;
        worker.onmessage = handler;
      };
      document.querySelector<HTMLInputElement>('input[aria-label="Formula bar"]')!.focus();
    }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
    await page.getByLabel('Formula bar', { exact: true }).fill('draft');
    await page.waitForFunction(() => !!(window as any).draftReadHeld);
    if (releaseReadFirst) await page.evaluate(() => (window as any).releaseDraftRead());
    await page.evaluate(async () => {
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A1')
        .setValues([[99]]);
    });
    if (!releaseReadFirst) await page.evaluate(() => (window as any).releaseDraftRead());
    await page.getByLabel('Formula bar', { exact: true }).press('Enter');
    await expect(
      page.getByRole('region', { name: 'Unconfirmed formula bar drafts' }).locator('textarea'),
    ).toHaveValue('draft');
    const submission = await page.evaluate(() => (window as any).draftSubmission);
    if (releaseReadFirst) {
      expect(submission.protocolVersion).toBeUndefined();
      expect(submission.payload.cells[0].expected.input).toEqual({ type: 'number', value: 42 });
    } else expect(submission.protocolVersion).toBe(1);
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
      ),
    ).toEqual([[99, 2]]);
  });
}

for (const surface of ['cell', 'formula'] as const) {
  for (const change of ['same-cell', 'sort'] as const) {
    test(`${surface} draft does not overwrite a concurrent ${change}`, async ({ page }) => {
      await page.goto('/');
      await expect(page.locator('#scene-select')).toBeEnabled();
      await page.evaluate(async (url) => {
        const { createWorkbookFile } = await import(/* @vite-ignore */ url);
        const app = window.opensheet;
        await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
        const sheet = app.getWorkbook().getSheets()[0];
        await sheet.range('A1:A2').setValues([[42], [10]]);
        app.selection.set({
          sheetId: sheet.id,
          startRow: 0,
          endRow: 1,
          startColumn: 0,
          endColumn: 1,
        });
        await app.getGrid()!.ready();
      }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
      if (surface === 'cell') {
        const grid = page.getByRole('grid', { exact: true });
        await grid.focus();
        await grid.press('Enter');
        await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('draft');
      } else await page.getByLabel('Formula bar', { exact: true }).fill('draft');
      await page.evaluate((change) => {
        const book = window.opensheet.getWorkbook();
        const worker = (book as any).worker as Worker;
        const post = worker.postMessage.bind(worker);
        const handler = worker.onmessage!;
        let firstId = -1;
        worker.postMessage = ((message: any, ...rest: any[]) => {
          if (message.method === 'execute' && firstId === -1) firstId = message.id;
          return Reflect.apply(post, worker, [message, ...rest]);
        }) as typeof worker.postMessage;
        worker.onmessage = (event) => {
          if (event.data.id === firstId) (window as any).contentHeld = event;
          else handler.call(worker, event);
        };
        (window as any).contentRelease = () => {
          handler.call(worker, (window as any).contentHeld);
          worker.postMessage = post;
          worker.onmessage = handler;
        };
        const sheet = book.getSheets()[0];
        (window as any).contentWrite =
          change === 'sort' ? sheet.sort('A1:A2', 0, 'asc') : sheet.range('A1').setValues([[99]]);
      }, change);
      await page.waitForFunction(() => !!(window as any).contentHeld);
      if (surface === 'cell')
        await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Enter');
      else await page.getByLabel('Formula bar', { exact: true }).press('Enter');
      const recovery = page.getByRole('region', {
        name: surface === 'cell' ? 'Unconfirmed cell drafts' : 'Unconfirmed formula bar drafts',
      });
      await expect(recovery.locator('textarea')).toHaveValue('draft');
      await page.evaluate(async () => {
        (window as any).contentRelease();
        await (window as any).contentWrite;
      });
      const values = await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
      );
      expect(values).toEqual(change === 'sort' ? [[10], [42]] : [[99], [10]]);
    });
  }
}

for (const surface of ['cell', 'formula'] as const) {
  test(`${surface} edit can follow an unrelated pending cell write`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    await page.evaluate(async (url) => {
      const { createWorkbookFile } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
      const sheet = app.getWorkbook().getSheets()[0];
      await sheet.range('A1:B1').setValues([[1, 2]]);
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
    }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
    if (surface === 'cell') {
      const grid = page.getByRole('grid', { exact: true });
      await grid.focus();
      await grid.press('Enter');
      await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('7');
    } else await page.getByLabel('Formula bar', { exact: true }).fill('7');
    const revision = await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let firstId = -1;
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute' && firstId === -1) firstId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === firstId) (window as any).unrelatedHeld = event;
        else handler.call(worker, event);
      };
      (window as any).unrelatedRelease = () => {
        handler.call(worker, (window as any).unrelatedHeld);
        worker.postMessage = post;
        worker.onmessage = handler;
      };
      (window as any).unrelatedWrite = book
        .getSheets()[0]
        .range('B1')
        .setValues([[9]]);
      return book.revision;
    });
    await page.waitForFunction(() => !!(window as any).unrelatedHeld);
    if (surface === 'cell')
      await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Enter');
    else await page.getByLabel('Formula bar', { exact: true }).press('Enter');
    await expect
      .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
      .toBe(revision + 2);
    await page.evaluate(async () => {
      (window as any).unrelatedRelease();
      await (window as any).unrelatedWrite;
    });
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
      ),
    ).toEqual([[7, 9]]);
  });
}

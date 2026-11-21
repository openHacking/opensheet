import { expect, test } from '@playwright/test';

test('first startup can retry after scene reference storage is restored', async ({ page }) => {
  await page.addInitScript(() => {
    const set = Storage.prototype.setItem;
    Storage.prototype.setItem = function (key, value) {
      if (this === localStorage && key === 'opensheet.scene.budget') {
        Storage.prototype.setItem = set;
        (window as any).failedInitialCandidate = value;
        throw new DOMException('Reference denied', 'QuotaExceededError');
      }
      return set.call(this, key, value);
    };
  });
  await page.goto('/#budget');
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.locator('#startup-recovery-message')).toContainText('Reference denied');
  const failed = await page.evaluate(() => ({
    candidate: (window as any).failedInitialCandidate as string,
    reference: localStorage.getItem('opensheet.scene.budget'),
  }));
  expect(failed.reference).toBeNull();
  expect(failed.candidate).toBeTruthy();
  await page.getByRole('button', { name: 'Try again' }).click();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).not.toBe(
    failed.candidate,
  );
});

test('a failed first scene reference keeps the current workbook and removes its candidate', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[765]]);
    const set = Storage.prototype.setItem;
    Storage.prototype.setItem = function (key, value) {
      if (this === localStorage && key === 'opensheet.scene.sales') {
        Storage.prototype.setItem = set;
        (window as any).failedSceneCandidate = value;
        throw new DOMException('Reference denied', 'QuotaExceededError');
      }
      return set.call(this, key, value);
    };
    return { id: book.id, revision: book.revision };
  });
  page.once('dialog', (dialog) => dialog.accept());
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#notice')).toContainText('Reference denied');
  expect(await page.locator('#scene-title').textContent()).toBe('Budget Simulator');
  const failed = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const candidate = (window as any).failedSceneCandidate as string;
    const request = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const lookup = db.transaction('heads').objectStore('heads').get(candidate);
    const head = await new Promise<unknown>((resolve, reject) => {
      lookup.onsuccess = () => resolve(lookup.result);
      lookup.onerror = () => reject(lookup.error);
    });
    db.close();
    return {
      id: book.id,
      revision: book.revision,
      value: await book.getSheets()[0].range('A1').getValues(),
      reference: localStorage.getItem('opensheet.scene.sales'),
      candidate,
      hasHead: head !== undefined,
    };
  });
  expect(failed).toEqual({
    ...initial,
    value: [[765]],
    reference: null,
    candidate: expect.any(String),
    hasHead: false,
  });
  page.once('dialog', (dialog) => dialog.accept());
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.sales'))).not.toBe(
    failed.candidate,
  );
});

test('a published scene survives failed attachment and can be reopened', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => {
    const app = window.opensheet;
    const attach = app.attachWorkbook.bind(app);
    app.attachWorkbook = async (...args) => {
      app.attachWorkbook = attach;
      throw new Error('Injected attachment failure');
    };
    return app.getWorkbook().id;
  });
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#notice')).toContainText('Injected attachment failure');
  const published = await page.evaluate(() => ({
    current: window.opensheet.getWorkbook().id,
    candidate: localStorage.getItem('opensheet.scene.sales'),
  }));
  expect(published.current).toBe(original);
  expect(published.candidate).toBeTruthy();
  await expect(page.locator('#notice')).toContainText(published.candidate!);
  await page.goto('/#sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(published.candidate);
});

test('an edit during first scene preparation prevents publishing a stale replacement', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const prototype = Object.getPrototypeOf(book);
    const ready = prototype.ready;
    prototype.ready = async function () {
      const result = await ready.call(this);
      if (this !== book) {
        prototype.ready = ready;
        await new Promise<void>((resolve) => ((window as any).resumeScene = resolve));
      }
      return result;
    };
    return book.id;
  });
  await page.locator('#scene-select').selectOption('sales');
  await expect.poll(() => page.evaluate(() => !!(window as any).resumeScene)).toBe(true);
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[987]]);
    (window as any).resumeScene();
  });
  await expect(page.locator('#notice')).toContainText('Workbook changed during scene creation');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.sales'))).toBeNull();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[987]]);
});

test('a draft started during scene preparation prevents publishing and remains editable', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const prototype = Object.getPrototypeOf(book);
    const ready = prototype.ready;
    prototype.ready = async function () {
      const result = await ready.call(this);
      if (this !== book) {
        prototype.ready = ready;
        await new Promise<void>((resolve) => ((window as any).resumeSceneDraft = resolve));
      }
      return result;
    };
  });
  await page.locator('#scene-select').selectOption('sales');
  await expect.poll(() => page.evaluate(() => !!(window as any).resumeSceneDraft)).toBe(true);
  const formula = page.getByRole('combobox', { name: 'Formula bar' });
  await formula.fill('unsaved draft');
  await page.evaluate(() => (window as any).resumeSceneDraft());
  await expect(page.locator('#notice')).toContainText('Workbook changed during scene creation');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.sales'))).toBeNull();
  await expect(formula).toHaveValue('unsaved draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
});

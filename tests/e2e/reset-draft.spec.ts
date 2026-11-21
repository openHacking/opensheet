import { expect, test } from '@playwright/test';

test('reset refuses an existing draft without creating or publishing a candidate', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  const formula = page.getByRole('combobox', { name: 'Formula bar' });
  await formula.fill('unfinished draft');
  await page.evaluate(() => document.querySelector<HTMLButtonElement>('#reset-demo')!.click());
  await expect(page.locator('#notice')).toContainText('Finish or cancel editing');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(original);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await expect(formula).toHaveValue('unfinished draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
});

test('a draft started during reset prevents candidate publication and remains editable', async ({
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
        (window as any).resetDraftCandidate = this.id;
        await new Promise<void>((resolve) => ((window as any).resumeResetDraft = resolve));
      }
      return result;
    };
    return book.id;
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect
    .poll(() => page.evaluate(() => Boolean((window as any).resumeResetDraft)))
    .toBe(true);
  const formula = page.getByRole('combobox', { name: 'Formula bar' });
  await formula.fill('unfinished reset draft');
  await page.evaluate(() => (window as any).resumeResetDraft());
  await expect(page.locator('#notice')).toContainText('Workbook changed during reset');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(original);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await expect(formula).toHaveValue('unfinished reset draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  const candidateAbsent = await page.evaluate(async () => {
    const request = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const read = db
      .transaction('heads')
      .objectStore('heads')
      .get((window as any).resetDraftCandidate);
    const result = await new Promise((resolve, reject) => {
      read.onsuccess = () => resolve(read.result);
      read.onerror = () => reject(read.error);
    });
    db.close();
    return result === undefined;
  });
  expect(candidateAbsent).toBe(true);
  await formula.press('Enter');
  await expect.poll(() => page.evaluate(() => window.opensheet.hasDraft)).toBe(false);
  await expect
    .poll(() =>
      page.evaluate(() => window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues()),
    )
    .toEqual([['unfinished reset draft']]);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['unfinished reset draft']]);
});

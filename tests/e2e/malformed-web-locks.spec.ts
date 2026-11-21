import { expect, test } from '@playwright/test';

for (const mode of ['malformed', 'not-allowed'] as const) {
  test(`Worker ${mode} Web Locks preserves the saved workbook for retry`, async ({ page }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const id = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[482]]);
      return book.id;
    });
    let inject = true;
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      const body = await response.text();
      const replacement =
        mode === 'malformed'
          ? '123'
          : "function() { return Promise.reject(new DOMException('Denied', 'NotAllowedError')); }";
      await route.fulfill({
        response,
        body: inject
          ? `${body}\nObject.getPrototypeOf(navigator.locks).request = ${replacement};`
          : body,
      });
    });
    await page.reload();
    await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await expect(page.locator('#startup-recovery-message')).toContainText('WEB_LOCKS_UNAVAILABLE');
    expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
    inject = false;
    await page.getByRole('button', { name: 'Try again' }).click();
    await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[482]]);
    expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
  });
}

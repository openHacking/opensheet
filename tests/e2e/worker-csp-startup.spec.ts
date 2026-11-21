import { expect, test } from '@playwright/test';

for (const failure of ['policy', 'missing-asset'] as const) {
  test(`startup reports ${failure} and saved data survives retry`, async ({ page }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[24680]]);
      return { id: book.id, revision: book.revision };
    });
    await page.addInitScript(() => {
      (window as any).workerPolicyViolations = [];
      document.addEventListener('securitypolicyviolation', (event) => {
        if (event.effectiveDirective === 'worker-src')
          (window as any).workerPolicyViolations.push(event.effectiveDirective);
      });
    });
    if (failure === 'policy') {
      await page.route('**/', async (route) => {
        const response = await route.fetch();
        await route.fulfill({
          response,
          headers: { ...response.headers(), 'content-security-policy': "worker-src 'none'" },
        });
      });
    } else {
      await page.route('**/engine.worker.js*', (route) => route.abort('failed'));
    }
    await page.reload();
    await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    const message = await page.locator('#startup-recovery-message').textContent();
    const violations = await page.evaluate(() => (window as any).workerPolicyViolations);
    await page.unroute(failure === 'policy' ? '**/' : '**/engine.worker.js*');
    await page.reload();
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const after = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        id: book.id,
        revision: book.revision,
        writer: !book.isReadOnly,
        value: await book.getSheets()[0].range('A1').getValues(),
      };
    });
    if (failure === 'policy') expect(violations).toContain('worker-src');
    else expect(violations).toEqual([]);
    expect(message).toContain('security policy');
    expect(message).toContain('Worker asset path');
    expect(after).toEqual({ ...before, writer: true, value: [[24680]] });
  });
}

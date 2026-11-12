import { expect, test } from '@playwright/test';
test('observe startup readiness', async ({ page }) => {
  test.setTimeout(45000);
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  const start = Date.now();
  await page.goto('/#budget');
  const navigation = Date.now() - start;
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator', { timeout: 30000 });
  console.log(JSON.stringify({ navigation, ready: Date.now() - start, errors }));
  expect(errors).toEqual([]);
});

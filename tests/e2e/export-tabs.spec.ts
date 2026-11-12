import { expect, test } from '@playwright/test';

test('export format tabs support keyboard navigation and expose the selected preview', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.getByRole('tab', { name: 'Code export' }).click();
  const tabs = page.getByRole('tablist', { name: 'Export format' });
  const latex = tabs.getByRole('tab', { name: 'LaTeX' });
  const markdown = tabs.getByRole('tab', { name: 'Markdown' });
  const tsv = tabs.getByRole('tab', { name: 'TSV' });
  await latex.focus();
  await page.keyboard.press('ArrowRight');
  await expect(markdown).toBeFocused();
  await expect(markdown).toHaveAttribute('aria-selected', 'true');
  await expect(latex).toHaveAttribute('tabindex', '-1');
  await expect(page.getByRole('tabpanel', { name: 'Markdown' })).toBeVisible();
  await page.keyboard.press('End');
  await expect(tsv).toBeFocused();
  await expect(page.getByRole('tabpanel', { name: 'TSV' })).toBeVisible();
  await page.keyboard.press('ArrowRight');
  await expect(latex).toBeFocused();
  await page.keyboard.press('Home');
  await expect(latex).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByLabel('Selected cells only')).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('tabpanel', { name: 'LaTeX' })).toBeFocused();
  await expect(tabs.getByLabel('Selected cells only')).toHaveCount(0);
});

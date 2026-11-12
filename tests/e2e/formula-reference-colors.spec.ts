import { expect, test } from '@playwright/test';

test('reference underlines remain visible at 200% zoom on a wide viewport', async ({ page }) => {
  await page.setViewportSize({ width: 1920, height: 1080 });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    document.body.style.zoom = '2';
  });
  const input = page.getByRole('combobox', { name: 'Formula bar' });
  await input.fill('=SUM(A1,B2,C3)');
  await expect(page.locator('.os-formula-reference-mirror')).toBeVisible();
  await expect(page.locator('.os-formula-reference-mirror b')).toHaveText(['A1', 'B2', 'C3']);
  const errors = await input.evaluate((element) => {
    if (!(element instanceof HTMLInputElement)) throw new Error('Expected formula input');
    const mirror = element.parentElement!.querySelector('.os-formula-reference-mirror')!;
    const plain = document.createElement('span');
    plain.textContent = element.value;
    Object.assign(plain.style, {
      position: 'absolute',
      whiteSpace: 'pre',
      font: getComputedStyle(element).font,
      padding: '0',
      fontKerning: 'none',
      fontVariantLigatures: 'none',
    });
    mirror.append(plain);
    const range = document.createRange();
    const origin = plain.getBoundingClientRect().left;
    const result = [...mirror.querySelectorAll('b')].map((mark) => {
      const start = element.value.indexOf(mark.textContent!);
      range.setStart(plain.firstChild!, 0);
      range.setEnd(plain.firstChild!, start);
      return Math.abs(
        mark.getBoundingClientRect().left -
          mirror.getBoundingClientRect().left -
          (range.getBoundingClientRect().right - origin),
      );
    });
    plain.remove();
    return result;
  });
  expect(errors).toHaveLength(3);
  for (const error of errors) expect(error).toBeLessThan(2);
});

test('formula references use the grid outline colors and clear on blur', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const input = page.getByRole('combobox', { name: 'Formula bar' });
  await input.fill('=SUM(A1,B2,"C3")');
  const marks = page.locator('.os-formula-reference-mirror b');
  await expect(marks).toHaveText(['A1', 'B2']);
  expect(
    await marks.evaluateAll((nodes) =>
      nodes.map((node) => (node as HTMLElement).style.textDecorationColor),
    ),
  ).toEqual(['rgb(37, 99, 235)', 'rgb(217, 119, 6)']);
  await input.press('Escape');
  await expect(marks).toHaveCount(0);
});

test('overflowing formula never displays misaligned underlines', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const input = page.getByRole('combobox', { name: 'Formula bar' });
  const formula =
    '=SUM(' + Array.from({ length: 50 }, (_, index) => `A${index + 1}`).join(',') + ')';
  await input.fill(formula);
  await input.press('End');
  await expect
    .poll(() => input.evaluate((element) => element.scrollWidth - element.clientWidth))
    .toBeGreaterThan(0);
  await expect(page.locator('.os-formula-reference-mirror')).toBeHidden();
  const styles = await input.evaluate((element) => {
    const mirror = element.parentElement!.querySelector('.os-formula-reference-mirror')!;
    const a = getComputedStyle(element),
      b = getComputedStyle(mirror);
    return {
      inputFont: a.fontFamily + '/' + a.fontSize + '/' + a.letterSpacing,
      mirrorFont: b.fontFamily + '/' + b.fontSize + '/' + b.letterSpacing,
    };
  });
  expect(styles.mirrorFont).toBe(styles.inputFont);
  await expect(page.locator('.os-formula-reference-mirror b')).toHaveCount(50);
  await input.fill('=A1+B2');
  await expect(page.locator('.os-formula-reference-mirror')).toBeVisible();
});

test('ignores unresolved references and hides stale underlines during composition', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const input = page.getByRole('combobox', { name: 'Formula bar' });
  await input.fill('=Missing!A1+A1+A1');
  const mirror = page.locator('.os-formula-reference-mirror');
  const marks = mirror.locator('b');
  await expect(marks).toHaveText(['A1', 'A1']);
  expect(
    await marks.evaluateAll((nodes) =>
      nodes.map((node) => (node as HTMLElement).style.textDecorationColor),
    ),
  ).toEqual(['rgb(37, 99, 235)', 'rgb(217, 119, 6)']);
  await input.dispatchEvent('compositionstart');
  await expect(mirror).toBeHidden();
  await input.dispatchEvent('compositionend');
  await expect(mirror).toBeVisible();
  await expect(marks).toHaveText(['A1', 'A1']);
});

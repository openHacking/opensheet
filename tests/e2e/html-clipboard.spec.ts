import { test, expect } from '@playwright/test';
import { readForPoll } from './read-for-poll.js';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:C3').clear({ all: true });
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  });
});

test('handles deep markup and exact text limits without counting a removed block newline', async ({
  page,
}) => {
  await page.evaluate(() =>
    window.opensheet
      .getGrid()!
      .pasteHTML(
        '<table><tr><td><p>' +
          'x'.repeat(32768) +
          '</p></td><td>' +
          '<span>'.repeat(12000) +
          'deep' +
          '</span>'.repeat(12000) +
          '</td><td>line<br></td></tr></table>',
      ),
  );
  await expect
    .poll(async () => {
      const values = await readForPoll(page, 'A1:C1');
      if (values === 'retry') return values;
      return [
        typeof values[0][0] === 'string' ? values[0][0].length : 0,
        values[0][1],
        values[0][2],
      ];
    })
    .toEqual([32768, 'deep', 'line\n']);
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.evaluate(() =>
    window.opensheet
      .getGrid()!
      .pasteHTML('<table><tr><td><p>' + 'x'.repeat(32769) + '</p></td></tr></table>'),
  );
  await expect(page.locator('#notice')).toContainText('Pasted cell text is too long');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
});

test('preserves block boundaries without trailing synthetic newlines or embedded foreign text', async ({
  page,
}) => {
  await page.evaluate(() =>
    window.opensheet
      .getGrid()!
      .pasteHTML(
        '<table><tr><td><div>first</div>second<p>third</p>end</td><td><div>42</div></td><td>safe<svg><script>hiddenScript</script><text>drawing</text></svg><math><mi>hiddenMath</mi></math><video>fallback</video>end</td></tr></table>',
      ),
  );
  await expect
    .poll(() => readForPoll(page, 'A1:C1'))
    .toEqual([['first\nsecond\nthird\nend', 42, 'safeend']]);
});

test('HTML paste event preserves cell text and supported formats in one history entry', async ({
  page,
}) => {
  await page.getByRole('grid', { exact: true }).evaluate((element) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', 'fallback must not win');
    clipboardData.setData(
      'text/html',
      '<table><tr><td style="font-weight:700;color:rgb(18,52,86);background-color:#abc;font-size:12pt;text-align:right">42</td><td>first<br>second &amp; third</td></tr><tr><td style="font-style:italic;text-decoration:underline;white-space:pre-wrap">=A1*2</td></tr></table>',
    );
    // Firefox does not retain constructor-supplied clipboardData on synthetic ClipboardEvents.
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    element.dispatchEvent(event);
  });
  await expect
    .poll(() => readForPoll(page, 'A1:B2'))
    .toEqual([
      [42, 'first\nsecond & third'],
      [84, null],
    ]);
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const first = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
    const second = book.getStyle((await book.getCell(sheet.id, 1, 0))?.styleId);
    await book.undo();
    const undone = await sheet.range('A1:B2').getValues();
    await book.redo();
    await app.open(book.id);
    return {
      first,
      second,
      undone,
      reopened: await app.getWorkbook().getSheets()[0].range('A1:B2').getValues(),
    };
  });
  expect(result.first).toEqual({
    bold: true,
    color: '#123456',
    background: '#aabbcc',
    fontSize: 16,
    align: 'right',
  });
  expect(result.second).toEqual({ italic: true, underline: true, wrap: true });
  expect(result.undone).toEqual([
    [null, null],
    [null, null],
  ]);
  expect(result.reopened).toEqual([
    [42, 'first\nsecond & third'],
    [84, null],
  ]);
});

test('clipboard markup remains inert and does not import script or embedded resource text', async ({
  page,
}) => {
  const requests: string[] = [];
  page.on('request', (request) => {
    if (request.url().includes('clipboard-probe')) requests.push(request.url());
  });
  await page.evaluate(() =>
    window.opensheet
      .getGrid()!
      .pasteHTML(
        '<table onclick="window.clipboardInjected=1"><tr><td>safe<script>window.clipboardInjected=1</script><style>body{display:none}</style><img src="/clipboard-probe" onerror="window.clipboardInjected=1"><iframe src="/clipboard-probe">hidden</iframe><a href="javascript:alert(1)"> link</a></td></tr></table>',
      ),
  );
  await expect.poll(() => readForPoll(page, 'A1')).toEqual([['safe link']]);
  expect(await page.evaluate(() => Object.hasOwn(window, 'clipboardInjected'))).toBe(false);
  expect(requests).toEqual([]);
  await expect(
    page.locator('table[onclick], img[src="/clipboard-probe"], iframe[src="/clipboard-probe"]'),
  ).toHaveCount(0);
});

test('rejects merged, nested and oversized HTML atomically and falls back for non-table HTML', async ({
  page,
}) => {
  for (const html of [
    '<table><tr><td colspan="2">merged</td></tr></table>',
    '<table><tr><td><table><tr><td>nested</td></tr></table></td></tr></table>',
    '<table><tr>' + '<td>x</td>'.repeat(1001) + '</tr></table>',
  ]) {
    const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
    await page.evaluate((html) => window.opensheet.getGrid()!.pasteHTML(html), html);
    await expect(page.locator('#notice')).toBeVisible();
    expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  }
  await page.evaluate(() => window.opensheet.getGrid()!.pasteHTML('<p>plain</p>', 'fallback\t2'));
  await expect.poll(() => readForPoll(page, 'A1:B1')).toEqual([['fallback', 2]]);
  await page.getByLabel('Read only', { exact: true }).check();
  await page.evaluate(() =>
    window.opensheet.getGrid()!.pasteHTML('<table><tr><td>forbidden</td></tr></table>'),
  );
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['fallback']]);
});

test('rejects empty rows that would expand a wide HTML table past the cell budget', async ({
  page,
}) => {
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.evaluate(() => {
    const html =
      '<table><tr>' + '<td>x</td>'.repeat(1000) + '</tr>' + '<tr></tr>'.repeat(200) + '</table>';
    window.opensheet.getGrid()!.pasteHTML(html);
  });
  await expect(page.locator('#notice')).toContainText('Paste dimensions are too large');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[null]]);
});

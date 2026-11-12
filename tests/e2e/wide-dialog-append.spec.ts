import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

for (const kind of ['sort', 'filter'] as const) {
  test(`${kind} appends wide-sheet conditions without rebuilding existing rows`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    await page.evaluate(
      async (fileJSON) => {
        const app = window.opensheet;
        await app.load(JSON.parse(fileJSON));
        const sheet = app.getWorkbook().getSheets()[0];
        app.selection.set({
          sheetId: sheet.id,
          startRow: 0,
          endRow: 2,
          startColumn: 0,
          endColumn: 1000,
        });
      },
      JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Wide', rows: 2, columns: 1000 }] })),
    );
    await page
      .getByLabel('Sheet actions')
      .selectOption(kind === 'sort' ? 'sort-custom' : 'filter-custom');
    const result = await page.evaluate((kind) => {
      const first = document.querySelector<HTMLSelectElement>(
        `[aria-label="${kind === 'sort' ? 'Level 1 column' : 'Condition 1 column'}"]`,
      )!;
      const container = first.closest('.sort-level')!.parentElement!;
      const button = [...document.querySelectorAll('button')].find(
        (element) => element.textContent === (kind === 'sort' ? 'Add sort level' : 'Add condition'),
      )!;
      const observer = new MutationObserver(() => {});
      observer.observe(container, { childList: true });
      const start = performance.now();
      for (let i = 0; i < 15; i++) button.click();
      const ms = performance.now() - start;
      const records = observer.takeRecords();
      const last = container.lastElementChild!.querySelector('select');
      const beforeRemoval = [
        ...container.querySelectorAll<HTMLSelectElement>('select[aria-label$="column"]'),
      ].map((select) => select.options.length);
      container.children[7].querySelector('button')!.click();
      const removalRecords = observer.takeRecords();
      observer.disconnect();
      return {
        ms,
        rows: container.childElementCount,
        removed: records.reduce((sum, record) => sum + record.removedNodes.length, 0),
        added: records.reduce((sum, record) => sum + record.addedNodes.length, 0),
        firstRetained: first === container.querySelector('select'),
        columns: beforeRemoval,
        removalAdded: removalRecords.reduce((sum, record) => sum + record.addedNodes.length, 0),
        removalRemoved: removalRecords.reduce((sum, record) => sum + record.removedNodes.length, 0),
        lastRetained: last === container.lastElementChild!.querySelector('select'),
        lastLabel: last?.getAttribute('aria-label'),
      };
    }, kind);
    expect(result.rows).toBe(15);
    expect(result.removed).toBe(0);
    expect(result.added).toBe(15);
    expect(result.firstRetained).toBe(true);
    expect(result.columns).toEqual(Array(16).fill(1000));
    expect(result.removalAdded).toBe(0);
    expect(result.removalRemoved).toBe(1);
    expect(result.lastRetained).toBe(true);
    expect(result.lastLabel).toBe(kind === 'sort' ? 'Level 15 column' : 'Condition 15 column');
    console.log(`${kind} 15 additions, 1000 columns: ${result.ms.toFixed(1)} ms`);
  });
}

import { expect, test } from '@playwright/test';

test('post-commit IndexedDB failure keeps the committed head and history recoverable', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let armPostCommit = false, failPostCommit = false;
addEventListener('message', e => { if (e.data.injectPostCommit) { armPostCommit = true; postMessage({faultStage: 'armed'}); } });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (armPostCommit && this.name === 'heads') {
    armPostCommit = false;
    failPostCommit = true;
    postMessage({faultStage: 'head'});
  }
  return originalPut.apply(this, args);
};
const originalGet = IDBObjectStore.prototype.get;
IDBObjectStore.prototype.get = function(...args) {
  if (failPostCommit && this.name === 'records') {
    failPostCommit = false;
    postMessage({faultStage: 'read'});
    throw new DOMException('Injected post-commit read failure', 'UnknownError');
  }
  return originalGet.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const values = await book.getSheets()[0].range('A1:B1').getValues();
    const revision = book.revision;
    const stages: string[] = [];
    (book as any).worker.addEventListener('message', (event: MessageEvent) => {
      if (event.data.faultStage) stages.push(event.data.faultStage);
    });
    (book as any).worker.postMessage({ injectPostCommit: true });
    let code = '';
    try {
      await book
        .getSheets()[0]
        .range('A1:B1')
        .setValues([[101, 202]]);
    } catch (error) {
      code = (error as { code?: string }).code ?? '';
    }
    return { id: book.id, revision, values, code, state: book.persistenceState.state, stages };
  });
  expect(before.stages).toContain('read');
  expect(before.code).toBe('WORKER_FAILED');
  expect(before.state).toBe('failed');
  await expect(page.locator('#save-state')).toContainText('Reopen workbook');
  const after = await page.evaluate(async (id) => {
    await window.opensheet.open(id);
    const book = window.opensheet.getWorkbook();
    const revision = book.revision;
    const values = await book.getSheets()[0].range('A1:B1').getValues();
    const undo = await book.undo();
    const undone = await book.getSheets()[0].range('A1:B1').getValues();
    const redo = await book.redo();
    const redone = await book.getSheets()[0].range('A1:B1').getValues();
    return { revision, values, undo, undone, redo, redone };
  }, before.id);
  expect(after).toEqual({
    revision: before.revision + 1,
    values: [[101, 202]],
    undo: true,
    undone: before.values,
    redo: true,
    redone: [[101, 202]],
  });
});

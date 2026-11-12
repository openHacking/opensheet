import { expect, test } from '@playwright/test';
import { resolve } from 'node:path';

test('invalid search arguments fail predictably without affecting data or later searches', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:A2').setValues([['needle'], ['other']]);
    const revision = book.revision;
    const rpc = (book as any).rpc;
    let requests = 0;
    (book as any).rpc = function (...args: unknown[]) {
      requests++;
      return rpc.apply(this, args);
    };
    const code = async (fn: () => unknown) => {
      try {
        await fn();
        return 'success';
      } catch (error) {
        return (error as { code?: string }).code ?? 'uncoded';
      }
    };
    const invalid = [
      await code(() => book.search(sheet.id, 42 as any)),
      await code(() => book.search(sheet.id, 'x'.repeat(65537))),
      await code(() => book.search(sheet.id, 'needle', -2)),
      await code(() => book.search(sheet.id, 'needle', Number.NaN)),
      await code(() => book.search(sheet.id, 'needle', Number.MAX_SAFE_INTEGER + 1)),
      await code(() => book.search(sheet.id, 'needle', -1, { direction: 'sideways' as any })),
      await code(() => book.search(sheet.id, 'needle', -1, null as any)),
      await code(() => book.search(sheet.id, 'needle', -1, { signal: {} as any })),
    ];
    const localRequests = requests;
    const unknownSheet = await code(() => book.search('missing-sheet', 'needle'));
    const found = await book.search(sheet.id, 'needle');
    const maxQuery = await book.search(sheet.id, 'x'.repeat(65536));
    let directionReads = 0;
    await book.search(sheet.id, 'needle', -1, {
      get direction() {
        directionReads++;
        return directionReads === 1 ? 'previous' : 'next';
      },
    });
    const last = await book.search(sheet.id, 'needle', Number.MAX_SAFE_INTEGER, {
      direction: 'previous',
    });
    (book as any).rpc = rpc;
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      invalid,
      localRequests,
      unknownSheet,
      found,
      maxQuery,
      directionReads,
      last,
      revision,
      after: book.revision,
      values: await sheet.range('A1:A2').getValues(),
    };
  });
  expect(result.invalid).toEqual([
    'INVALID_ARGUMENT',
    'LIMIT_EXCEEDED',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
  ]);
  expect(result.localRequests).toBe(0);
  expect(result.unknownSheet).toBe('INVALID_ARGUMENT');
  expect(result.found).toEqual({ count: 1, match: { row: 0, column: 0 } });
  expect(result.maxQuery).toEqual({ count: 0, match: undefined });
  expect(result.directionReads).toBe(1);
  expect(result.last).toEqual(result.found);
  expect(result.after).toBe(result.revision);
  expect(result.values).toEqual([['needle'], ['other']]);
});

test('Worker independently rejects malformed search requests', async ({ page }) => {
  const core = (name: string) => `/@fs/${resolve(`packages/core/src/${name}.ts`)}`;
  await page.route('**/search-arguments.worker.js*', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { Engine } from ${JSON.stringify(core('engine'))};
      import { createSnapshot } from ${JSON.stringify(core('model'))};
      import { Database } from ${JSON.stringify(core('storage'))};
      (async()=>{
        const snapshot=createSnapshot({sheets:[{name:'Data',rows:2,columns:2}]});
        const sheet=snapshot.sheets[0], db=await Database.open('search-args-'+crypto.randomUUID());
        const engine=new Engine(db,snapshot.workbookId,0);
        await engine.open(snapshot);
        const revision=engine.view().revision;
        const code=async(args)=>{ try { await engine.search(...args); return 'success'; } catch(e) { return e.code ?? 'uncoded'; } };
        const invalid=[];
        for(const args of [[sheet.id,42,-1,false,'next'],[sheet.id,'x'.repeat(65537),-1,false,'next'],[sheet.id,'x',NaN,false,'next'],[sheet.id,'x',-1,'yes','next'],[sheet.id,'x',-1,false,'other'],['missing','x',-1,false,'next']]) invalid.push(await code(args));
        const found=await engine.search(sheet.id,'x');
        const after=engine.view().revision;
        await engine.clear(); engine.close();
        self.postMessage({invalid,found,revision,after});
      })().catch(error=>self.postMessage({error:String(error)}));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolveResult, reject) => {
        const worker = new Worker('/search-arguments.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker timed out'));
        }, 15000);
        worker.onerror = (event) => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error(event.message));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timer);
          worker.terminate();
          resolveResult(data);
        };
      }),
  );
  expect(result.invalid).toEqual([
    'INVALID_ARGUMENT',
    'LIMIT_EXCEEDED',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
    'INVALID_ARGUMENT',
  ]);
  expect(result.found).toEqual({ count: 0, match: undefined });
  expect(result.after).toBe(result.revision);
});

import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('file reader iterates a large sparse range in bounded batches', async ({ page }) => {
  const modulePath = `/@fs${fileURLToPath(new URL('../../packages/core/src/index.ts', import.meta.url))}`;
  await page.route('**/file-reader-batches.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { createWorkbookFile, encodeFileCells, FileReader } from ${JSON.stringify(modulePath)};
      try {
        const file = createWorkbookFile({sheets:[{name:'Data',rows:50000,columns:1}]});
        const sheet = file.sheets[0];
        sheet.blocks = [0,49999].map((row)=>({row:Math.floor(row/64),column:0,data:encodeFileCells([{rowId:row,columnId:0,input:{type:'number',value:row+1}}],new Map())}));
        const book = new FileReader(file);
        try {
          let total = 0, max = 0, sum = 0, batches = 0;
          for (const batch of book.getSheetById(sheet.id).range('A1:A50000').batches({maxCells:1024})) {
            batches++; max = Math.max(max,batch.values.length); total += batch.values.length;
            for (const row of batch.values) sum += row[0] ?? 0;
          }
          self.postMessage({total,max,sum,batches});
        } finally { book.dispose(); }
      } catch (error) { self.postMessage({error:String(error)}); }`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/file-reader-batches.worker.js', { type: 'module' });
        worker.onmessage = ({ data }) => {
          worker.terminate();
          resolve(data);
        };
        worker.onerror = (event) => {
          worker.terminate();
          reject(new Error(event.message));
        };
      }),
  );
  expect(result).toEqual({ total: 50000, max: 1024, sum: 50001, batches: 49 });
});

test('mutating returned formula errors does not corrupt subsequent worker reads', async ({
  page,
}) => {
  const modulePath = `/@fs${fileURLToPath(new URL('../../packages/core/src/index.ts', import.meta.url))}`;
  await page.route('**/file-reader-errors.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { createWorkbookFile, encodeFileCells, FileReader } from ${JSON.stringify(modulePath)};
      try {
        const file = createWorkbookFile({sheets:[{name:'Data',rows:2,columns:1}]});
        const sheet = file.sheets[0];
        file.styles.push({numberFormat:'0.00'});
        sheet.blocks = [{row:0,column:0,data:encodeFileCells([
          {rowId:0,columnId:0,input:{type:'formula',expression:'1/0'},styleId:'s0',link:{target:'https://example.com'}},
          {rowId:1,columnId:0,input:{type:'formula',expression:'A1+1'}}
        ],new Map([['s0',0]]))}];
        const book = new FileReader(file);
        try {
          const cell = book.getCell(sheet.id,0,0);
          cell.input.expression = '99';
          cell.link.target = 'https://changed.example';
          book.getStyle('s0').numberFormat = 'changed';
          [...book.cells(sheet.id)][1].input.expression = '99';
          book.value(sheet.id,0,0).error = 'changed';
          const range = book.getSheetById(sheet.id).range('A1:A2');
          const first = [...range.batches()][0].values;
          const before = first.map(row => row[0].error);
          first[1][0].error = 'changed';
          const after = range.getValues().map(row => row[0].error);
          range.bounds.endRow = 1000000000;
          let boundsError;
          try { range.getValues(); } catch (error) { boundsError = error.code; }
          self.postMessage({before,after,link:book.getCell(sheet.id,0,0).link.target,format:book.getStyle('s0').numberFormat,boundsError});
        } finally { book.dispose(); }
      } catch (error) { self.postMessage({error:String(error)}); }`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/file-reader-errors.worker.js', { type: 'module' });
        worker.onmessage = ({ data }) => {
          worker.terminate();
          resolve(data);
        };
        worker.onerror = (event) => {
          worker.terminate();
          reject(new Error(event.message));
        };
      }),
  );
  expect(result).toEqual({
    before: ['#DIV/0!', '#DIV/0!'],
    after: ['#DIV/0!', '#DIV/0!'],
    link: 'https://example.com',
    format: '0.00',
    boundsError: 'LIMIT_EXCEEDED',
  });
});

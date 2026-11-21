import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { readFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { chromium, firefox, webkit } from '@playwright/test';

// Execute the exact lazy outputs measured by bundle-breakdown.mjs on a localhost subpath.
export async function runLazyRuntime(temp, outputs, xlsxOutputs) {
  const html = `<!doctype html><html><head><link rel="stylesheet" href="./lazy.css"></head><body><main id="root" class="os-framework-host"></main><button id="convert">Convert XLSX</button><script type="module" src="./runtime.js"></script></body></html>`;
  const runtime = `
import './lazy.js';
try {
  const app = window.api.createOpenSheet({container:'#root'});
  const book = await app.createWorkbook({sheets:[{name:'Data',rows:5,columns:4}]});
  await book.getSheets()[0].range('A1:B1').setValues([[6,7]]);
  await book.getSheets()[0].range('C1').setFormulas([['A1+B1']]);
  await app.ready();
  window.demo = {app,book};
  document.querySelector('#convert').onclick = async () => {
    try {
      const {fromSheetJS,toSheetJS,XLSX} = await window.api.loadXlsx();
      const source = await window.demo.book.toJSON();
      const converted = toSheetJS(source);
      const bytes = XLSX.write(converted.workbook,{type:'array',bookType:'xlsx'});
      const parsed = XLSX.read(bytes,{type:'array',sheetStubs:true,bookFiles:true});
      const imported = fromSheetJS(parsed);
      await app.load(imported.file);
      let next = app.getWorkbook();
      const values = await next.getSheets()[0].range('A1:C1').getValues();
      const id = next.id;
      await app.open(id);
      next = app.getWorkbook();
      const reopened = await next.getSheets()[0].range('A1:C1').getValues();
      window.demo.book = next;
      window.demo.result = {values,reopened,bytes:bytes.byteLength,readOnly:next.isReadOnly};
      window.demo.runs = (window.demo.runs || 0) + 1;
    } catch (error) { document.body.dataset.error = String(error?.stack ?? error); }
  };
} catch(error) { document.body.dataset.error = String(error?.stack ?? error); }
`;
  const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
  const assets = new Map([
    ['index.html', Buffer.from(html)],
    ['runtime.js', Buffer.from(runtime)],
    ['engine.worker.js', readFileSync(join(temp, 'out/worker/engine.worker.js'))],
  ]);
  for (const [file, output] of Object.entries(outputs)) {
    const name = basename(file),
      bytes = readFileSync(join(temp, file));
    assert.ok(!assets.has(name), `Duplicate runtime asset name: ${name}`);
    assert.equal(digest(bytes), output.sha256, `Measured and served bytes differ: ${name}`);
    assets.set(name, bytes);
  }
  const prefix = '/embedded/sheets/';
  const server = createServer((req, res) => {
    const path = new URL(req.url, 'http://localhost').pathname;
    const file =
      path === prefix ? 'index.html' : path.startsWith(prefix) ? path.slice(prefix.length) : '';
    const bytes = assets.get(file);
    if (!bytes) {
      res.writeHead(404);
      res.end();
      return;
    }
    res.writeHead(200, {
      'Content-Type': file.endsWith('.css')
        ? 'text/css'
        : file.endsWith('.html')
          ? 'text/html'
          : 'text/javascript',
      'Cache-Control': 'no-store',
    });
    res.end(bytes);
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const results = [];
  try {
    for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
      const browser = await launcher.launch();
      try {
        const page = await browser.newPage(),
          requests = [],
          errors = [],
          failures = [];
        page.on('request', (request) => requests.push(new URL(request.url()).pathname));
        page.on('pageerror', (error) => errors.push(String(error)));
        page.on('response', (response) => {
          if (response.status() >= 400) failures.push(`${response.status()} ${response.url()}`);
        });
        await page.goto(`http://127.0.0.1:${server.address().port}${prefix}`);
        await page.waitForFunction(() => window.demo || document.body.dataset.error, null, {
          timeout: 30000,
        });
        const before = await page.evaluate(async () => {
          if (document.body.dataset.error) throw Error(document.body.dataset.error);
          return {
            values: await window.demo.book.getSheets()[0].range('A1:C1').getValues(),
            height: document.querySelector('.opensheet').getBoundingClientRect().height,
            stylesheets: document.querySelectorAll('link[rel=stylesheet]').length,
          };
        });
        const beforeRequests = [...requests];
        const xlsxPaths = xlsxOutputs.map((file) => prefix + basename(file));
        assert.deepEqual(before.values, [[6, 7, 13]]);
        assert.equal(before.height, 500);
        assert.equal(before.stylesheets, 1);
        assert.ok(beforeRequests.includes(prefix + 'engine.worker.js'));
        assert.ok(
          xlsxPaths.every((path) => !beforeRequests.includes(path)),
          'XLSX loaded before activation',
        );
        for (let run = 1; run <= 2; run++) {
          await page.locator('#convert').click();
          await page.waitForFunction(
            (run) => window.demo?.runs >= run || document.body.dataset.error,
            run,
            { timeout: 30000 },
          );
          const converted = await page.evaluate(() => {
            if (document.body.dataset.error) throw Error(document.body.dataset.error);
            return window.demo.result;
          });
          assert.deepEqual(converted.values, [[6, 7, 13]]);
          assert.deepEqual(converted.reopened, converted.values);
          assert.equal(converted.readOnly, false);
          assert.ok(converted.bytes > 0);
        }
        for (const path of xlsxPaths)
          assert.equal(requests.filter((request) => request === path).length, 1);
        assert.deepEqual(errors, []);
        assert.deepEqual(failures, []);
        results.push({
          engine,
          browser: browser.version(),
          before,
          beforeRequests,
          afterRequests: requests,
          converted: await page.evaluate(() => window.demo.result),
          errors,
          failures,
        });
      } finally {
        await browser.close();
      }
    }
    return {
      prefix,
      htmlSha256: digest(html),
      runtimeSha256: digest(runtime),
      workerSha256: digest(assets.get('engine.worker.js')),
      results,
    };
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
}

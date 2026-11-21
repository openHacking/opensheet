import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { readFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { chromium, firefox, webkit } from '@playwright/test';

// Execute the exact lazy outputs measured by bundle-breakdown.mjs on a localhost subpath.
export async function runLazyRuntime(temp, outputs, xlsxOutputs, { strictCsp = false } = {}) {
  const html = `<!doctype html><html><head><link rel="stylesheet" href="./lazy.css"></head><body><main id="root" class="os-framework-host"></main><button id="convert">Convert XLSX</button>${strictCsp ? '<script src="./probe.js"></script>' : ''}<script type="module" src="./runtime.js"></script></body></html>`;
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
    [
      'probe.js',
      Buffer.from(
        "try { new Function('return 1')(); document.body.dataset.cspBlocked = 'false'; } catch { document.body.dataset.cspBlocked = 'true'; }",
      ),
    ],
    [
      'probe.worker.js',
      Buffer.from(
        "self.onmessage = () => { try { new Function('return 1')(); self.postMessage(false); } catch { self.postMessage(true); } };",
      ),
    ],
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
  const csp = strictCsp
    ? "default-src 'none'; script-src 'self'; worker-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; font-src 'self'; base-uri 'none'; object-src 'none'"
    : null;
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
      ...(csp ? { 'Content-Security-Policy': csp } : {}),
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
        if (strictCsp)
          await page.addInitScript(() => {
            window.cspEvents = [];
            document.addEventListener('securitypolicyviolation', (event) =>
              window.cspEvents.push({
                directive: event.effectiveDirective,
                blocked: event.blockedURI,
                source: event.sourceFile,
              }),
            );
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
        let cspControls = null;
        if (strictCsp) {
          cspControls = await page.evaluate(async () => ({
            pageBlocked: document.body.dataset.cspBlocked === 'true',
            workerBlocked: await new Promise((resolve, reject) => {
              const worker = new Worker('./probe.worker.js', { type: 'module' });
              const timer = setTimeout(() => {
                worker.terminate();
                reject(new Error('CSP Worker probe timed out'));
              }, 5000);
              worker.onmessage = (event) => {
                clearTimeout(timer);
                worker.terminate();
                resolve(event.data);
              };
              worker.onerror = (event) => {
                clearTimeout(timer);
                worker.terminate();
                reject(new Error(event.message));
              };
              worker.postMessage(null);
            }),
          }));
          assert.deepEqual(cspControls, { pageBlocked: true, workerBlocked: true });
        }
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
        const cspEvents = strictCsp ? await page.evaluate(() => window.cspEvents) : [];
        for (const event of cspEvents)
          assert.ok(
            event.blocked === 'eval' && event.source.endsWith('/probe.js'),
            `Unexpected CSP violation: ${JSON.stringify(event)}`,
          );
        results.push({
          engine,
          browser: browser.version(),
          before,
          cspControls,
          cspEvents,
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
      csp,
      htmlSha256: digest(html),
      runtimeSha256: digest(runtime),
      pageProbeSha256: digest(assets.get('probe.js')),
      workerProbeSha256: digest(assets.get('probe.worker.js')),
      workerSha256: digest(assets.get('engine.worker.js')),
      results,
    };
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
}

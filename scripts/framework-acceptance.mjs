import { execFileSync, spawn } from 'node:child_process';
import {
  mkdtempSync,
  mkdirSync,
  readFileSync,
  writeFileSync,
  copyFileSync,
  existsSync,
  rmSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve, join, dirname } from 'node:path';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';
const option = (name, fallback) =>
  process.argv.find((arg) => arg.startsWith(`--${name}=`))?.slice(name.length + 3) ?? fallback;
const output = resolve(option('output', 'test-results/framework-acceptance.json'));
const artifactDirectory = option('artifacts', '');
const artifactManifest = artifactDirectory
  ? JSON.parse(readFileSync(join(resolve(artifactDirectory), 'manifest.json'), 'utf8'))
  : undefined;
if (existsSync(output)) throw new Error('Refusing to overwrite acceptance evidence');
const root = process.cwd(),
  consumer = mkdtempSync(join(tmpdir(), 'opensheet-framework-'));
const servers = [],
  hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  artifactDirectory: artifactDirectory ? resolve(artifactDirectory) : null,
  packages: [],
  results: [],
  scope:
    'Installs locally packed package tarballs in an isolated consumer with no OpenSheet source aliases; explicit Zod CSP configuration shim. Development Strict Mode and production Vite builds. Exact direct runtime versions; generated consumer dependency lock is preserved. Browser lifecycle assertions, not a memory benchmark.',
};
mkdirSync(dirname(output), { recursive: true });
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const run = (args, cwd = consumer, timeout = 120000) =>
  execFileSync('pnpm', args, {
    cwd,
    encoding: 'utf8',
    timeout,
    maxBuffer: 4 * 1024 * 1024,
  });
const readyServer = async (args, port) => {
  const server = spawn(
    process.execPath,
    [
      join(consumer, 'node_modules/vite/bin/vite.js'),
      ...args,
      '--host',
      '127.0.0.1',
      '--port',
      String(port),
      '--strictPort',
    ],
    { cwd: consumer, stdio: ['ignore', 'pipe', 'pipe'] },
  );
  servers.push(server);
  let log = '';
  server.stdout.on('data', (chunk) => {
    log += chunk;
  });
  server.stderr.on('data', (chunk) => {
    log += chunk;
  });
  const url = `http://127.0.0.1:${port}/embedded/sheets/`;
  for (let i = 0; i < 300; i++) {
    if (server.exitCode !== null) throw new Error(`Host server exited: ${log}`);
    if (log.includes(`:${port}/`)) {
      try {
        if ((await fetch(url)).ok) return url;
      } catch {}
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`Host server timeout: ${log}`);
};
try {
  mkdirSync(join(consumer, 'tarballs'));
  const dependencies = {
      react: '19.2.0',
      'react-dom': '19.2.0',
      vue: '3.5.42',
      vite: '7.3.7',
      zod: '4.6.6',
    },
    overrides = {};
  for (const folder of ['formula', 'core', 'renderer', 'plugin-sdk', 'opensheet', 'react', 'vue']) {
    const manifest = JSON.parse(
      readFileSync(join(root, 'packages', folder, 'package.json'), 'utf8'),
    );
    const archive = join(consumer, 'tarballs', `${folder}.tgz`);
    if (artifactManifest) {
      copyFileSync(join(resolve(artifactDirectory), `${folder}.tgz`), archive);
      assert.equal(
        hash(readFileSync(archive)),
        artifactManifest.packages.find((item) => item.name === manifest.name)?.sha256,
        `${folder}: candidate archive mismatch`,
      );
    } else run(['pack', '--out', archive], join(root, 'packages', folder));
    dependencies[manifest.name] = overrides[manifest.name] = `file:./tarballs/${folder}.tgz`;
    report.packages.push({
      name: manifest.name,
      version: manifest.version,
      sha256: hash(readFileSync(archive)),
    });
  }
  writeFileSync(
    join(consumer, 'package.json'),
    JSON.stringify(
      {
        name: 'opensheet-framework-host',
        private: true,
        type: 'module',
        dependencies,
        pnpm: { overrides },
      },
      null,
      2,
    ),
  );
  for (const file of ['index.html', 'main.js', 'vite.config.js', 'ssr.mjs', 'csp-zod.js'])
    copyFileSync(join(root, 'examples/framework-host', file), join(consumer, file));
  report.fixtureSha256 = hash(readFileSync(join(consumer, 'main.js')));
  report.configSha256 = hash(readFileSync(join(consumer, 'vite.config.js')));
  report.cspShimSha256 = hash(readFileSync(join(consumer, 'csp-zod.js')));
  report.ssrFixtureSha256 = hash(readFileSync(join(consumer, 'ssr.mjs')));
  process.stdout.write('Installing packed framework consumer…\n');
  run(
    ['install', '--prefer-offline', '--ignore-scripts', '--strict-peer-dependencies'],
    consumer,
    300000,
  );
  const lock = readFileSync(join(consumer, 'pnpm-lock.yaml'));
  const lockPath = output.replace(/\.json$/, '') + '.consumer-lock.yaml';
  writeFileSync(lockPath, lock, { flag: 'wx' });
  report.lock = { file: lockPath, sha256: hash(lock) };
  report.ssr = JSON.parse(run(['exec', 'node', 'ssr.mjs']));
  const template = readFileSync(join(consumer, 'index.html'), 'utf8');
  for (const framework of ['react', 'vue'])
    writeFileSync(
      join(consumer, `${framework}.html`),
      template.replace(
        '<main id="root"></main>',
        `<main id="root" data-ssr>${report.ssr[framework]}</main>`,
      ),
    );
  process.stdout.write('Building package-installed production host…\n');
  run(['exec', 'vite', 'build']);
  save();
  const urls = {
    development: await readyServer([], 5191),
    production: await readyServer(['preview'], 5192),
  };
  urls.hydration = urls.production;
  const launchers = { chromium, webkit, firefox };
  for (const engine of option('engines', 'chromium,webkit,firefox').split(',')) {
    const browser = await launchers[engine].launch();
    try {
      for (const [mode, url] of Object.entries(urls))
        for (const framework of ['react', 'vue']) {
          const context = await browser.newContext(),
            page = await context.newPage(),
            pageErrors = [];
          page.on('pageerror', (error) => pageErrors.push(String(error)));
          const consoleErrors = [];
          page.on('console', (message) => {
            if (message.type() === 'error') consoleErrors.push(message.text());
          });
          await page.addInitScript(() => {
            window.cspViolations = [];
            document.addEventListener('securitypolicyviolation', (event) =>
              window.cspViolations.push({
                directive: event.effectiveDirective,
                blocked: event.blockedURI,
              }),
            );
          });
          const result = { engine, browserVersion: browser.version(), mode, framework, checks: [] };
          report.results.push(result);
          save();
          try {
            const response = await page.goto(
              `${url}${mode === 'hydration' ? `${framework}.html` : ''}?framework=${framework}`,
            );
            if (mode !== 'development') {
              result.csp = response.headers()['content-security-policy'];
              assert.ok(result.csp?.includes("worker-src 'self'"));
              assert.ok(!result.csp.includes('unsafe-'));
            }
            await page.waitForFunction(() => !!window.hostFixture);
            const invoke = (method, arg) =>
              page.evaluate(({ method, arg }) => window.hostFixture[method](arg), { method, arg });
            const snapshot = () =>
              page.evaluate(() => ({
                ready: window.hostFixture.state.ready,
                changes: window.hostFixture.state.changes,
                errors: window.hostFixture.state.errors,
                ...window.hostFixture.resources(),
              }));
            await invoke('mount', { value: 1 });
            await page.waitForFunction(
              () =>
                window.hostFixture.state.ready === 1 || window.hostFixture.state.errors.length > 0,
            );
            assert.equal((await snapshot()).ready, 1, 'Host initialization failed');
            assert.equal((await snapshot()).workers, 1);
            assert.equal(
              await page
                .locator('#root > .os-framework-host')
                .evaluate((element) => getComputedStyle(element).height),
              '500px',
            );
            if (mode === 'hydration') {
              assert.equal(
                await page.evaluate(() => window.hostFixture.state.hydratedHostPreserved),
                true,
              );
              assert.deepEqual(
                await page.evaluate(() => window.hostFixture.state.hydrationErrors),
                [],
              );
              assert.deepEqual(consoleErrors, []);
              result.checks.push(
                'hydrates the server-rendered host without replacement or mismatch',
              );
            }
            if (framework === 'react' && mode === 'development')
              assert.ok(
                (await snapshot()).createdWorkers >= 2,
                'Strict Mode effect replay not observed',
              );
            await invoke('update', { value: 2, readOnly: true, version: 2 });
            const readOnly = await page.evaluate(async () => {
              const book = window.hostFixture.state.app.getWorkbook();
              let code;
              try {
                await book
                  .getSheets()[0]
                  .range('A1')
                  .setValues([[9]]);
              } catch (error) {
                code = error.code;
              }
              return { code, value: await book.getSheets()[0].range('A1').getValues() };
            });
            assert.deepEqual(readOnly, { code: 'READ_ONLY', value: [[1]] });
            result.checks.push('initialFile is mount-only; readOnly updates');
            await invoke('update', { readOnly: false });
            await page.evaluate(() =>
              window.hostFixture.state.app
                .getWorkbook()
                .getSheets()[0]
                .range('A1')
                .setValues([[7]]),
            );
            assert.equal((await snapshot()).changes.at(-1).version, 2);
            await invoke('replace', 3);
            const id = await page.evaluate(async () => {
              const book = window.hostFixture.state.app.getWorkbook();
              if ((await book.getSheets()[0].range('A1').getValues())[0][0] !== 3)
                throw new Error('Replacement failed');
              await book
                .getSheets()[0]
                .range('A1')
                .setValues([[4]]);
              return book.id;
            });
            assert.equal((await snapshot()).changes.length, 2);
            await invoke('unmount');
            assert.deepEqual(await invoke('reopen', id), { readOnly: false, value: [[4]] });
            assert.equal((await snapshot()).workers, 0);
            assert.equal((await snapshot()).grids, 0);
            result.checks.push(
              'current callbacks, controlled replacement, unmount and writer lock release',
            );
            await invoke('hold');
            await invoke('mount');
            await page.waitForFunction(() => window.hostFixture.state.pending > 0);
            await invoke('unmount');
            await invoke('release');
            assert.equal((await snapshot()).ready, 0);
            assert.equal((await snapshot()).errors.length, 0);
            assert.equal((await snapshot()).workers, 0);
            result.checks.push('unmount during async readiness suppresses stale callbacks');
            for (let navigation = 0; navigation < 10; navigation++) {
              await invoke('hold');
              await invoke('mount', { value: navigation });
              await page.waitForFunction(() => window.hostFixture.state.pending > 0);
              await invoke('unmount');
              await invoke('release');
              const after = await snapshot();
              assert.equal(after.ready, 0);
              assert.deepEqual(after.errors, []);
              assert.equal(after.workers, 0);
              assert.equal(after.grids, 0);
            }
            result.checks.push('ten rapid navigation cycles leave no stale readiness or Workers');
            await invoke('mount', { invalid: true });
            await page.waitForFunction(() => window.hostFixture.state.errors.length > 0);
            assert.equal((await snapshot()).ready, 0);
            await invoke('unmount');
            assert.equal((await snapshot()).workers, 0);
            result.checks.push('async initialization errors reach host');
            for (const phase of [
              'toolbar-observe',
              'toolbar-listener',
              'toolbar-disconnect',
              'formula-observe',
              'formula-listener',
              'formula-disconnect',
              'formula-dom',
              'editor-dom',
              'host-append',
            ]) {
              await invoke('failEditorConstruction', phase);
              await invoke('mount');
              await page.waitForFunction(() => window.hostFixture.state.errors.length > 0);
              assert.match(
                (await snapshot()).errors[0].message,
                /Injected editor construction failure/,
              );
              assert.equal((await snapshot()).ready, 0);
              assert.equal((await snapshot()).workers, 0);
              if (phase.endsWith('-disconnect'))
                assert.ok(
                  (await page.evaluate(() => window.hostFixture.reportedObserverErrors)).some(
                    (message) => message.includes('Injected observer cleanup failure'),
                  ),
                );
              await invoke('unmount');
              assert.deepEqual(
                await invoke('constructionResources'),
                { listeners: 0, observers: 0 },
                phase,
              );
              await invoke('restoreConstruction');
              await invoke('mount');
              await page.waitForFunction(() => window.hostFixture.state.ready === 1);
              await invoke('unmount');
              assert.equal((await snapshot()).workers, 0);
              assert.equal((await snapshot()).grids, 0);
            }
            result.checks.push('nine construction failures release resources and permit remount');
            if (framework === 'react') {
              for (const observerFailure of ['sync', 'async']) {
                await invoke('failEditorConstruction');
                await invoke('mount', { observerFailure });
                await page.waitForFunction(() => window.hostFixture.state.errors.length > 0);
                await page.waitForTimeout(50);
                assert.equal(
                  await page.evaluate(() => window.hostFixture.state.observerReceiverPreserved),
                  true,
                );
                assert.ok(
                  (await page.evaluate(() => window.hostFixture.reportedObserverErrors)).some(
                    (message) =>
                      message.includes(`Injected ${observerFailure} host observer failure`),
                  ),
                );
                await invoke('unmount');
                assert.deepEqual(await invoke('constructionResources'), {
                  listeners: 0,
                  observers: 0,
                });
                await invoke('restoreConstruction');
              }
              result.checks.push('throwing and rejecting initialization observers are contained');
            }
            await invoke('mount');
            await page.waitForFunction(() => window.hostFixture.state.ready === 1);
            await invoke('installLateAction');
            if (!(await page.getByRole('button', { name: 'Late action', exact: true }).isVisible()))
              await page.getByRole('button', { name: 'More tools', exact: true }).click();
            await page.getByRole('button', { name: 'Late action', exact: true }).click();
            await invoke('rejectLate');
            await page.waitForFunction(() => window.hostFixture.state.errors.length === 1);
            assert.match((await snapshot()).errors[0].message, /late host error/);
            await page.evaluate(() => {
              window.hostFixture.state.errors = [];
            });
            if (!(await page.getByRole('button', { name: 'Late action', exact: true }).isVisible()))
              await page.getByRole('button', { name: 'More tools', exact: true }).click();
            await page.getByRole('button', { name: 'Late action', exact: true }).click();
            await invoke('unmount');
            await invoke('rejectLate');
            await page.waitForTimeout(100);
            assert.equal(
              (await snapshot()).errors.length,
              0,
              'Host received an error after unmount',
            );
            assert.equal((await snapshot()).workers, 0);
            assert.equal((await snapshot()).grids, 0);
            result.checks.push('pending action rejection after unmount is silent');
            assert.deepEqual(pageErrors, []);
            const deployment = await page.evaluate(() => ({
              violations: window.cspViolations,
              styles: [...document.querySelectorAll('link[rel="stylesheet"]')].map((el) => el.href),
              workers: window.hostFixture.state.workerUrls,
            }));
            assert.deepEqual(deployment.violations, []);
            if (mode !== 'development') {
              assert.equal(deployment.styles.length, 1);
              for (const asset of [...deployment.styles, ...deployment.workers])
                assert.ok(new URL(asset).pathname.startsWith('/embedded/sheets/assets/'));
            }
            result.deployment = deployment;
            result.status = 'passed';
            save();
            process.stdout.write(`${engine} ${mode} ${framework}: passed\n`);
          } catch (error) {
            result.status = 'failed';
            result.error = String(error);
            result.pageErrors = pageErrors;
            result.observation = await page
              .evaluate(() =>
                window.hostFixture
                  ? {
                      ready: window.hostFixture.state.ready,
                      errors: window.hostFixture.state.errors,
                      workerUrls: window.hostFixture.state.workerUrls,
                      ...window.hostFixture.resources(),
                    }
                  : null,
              )
              .catch(() => null);
            save();
            throw error;
          } finally {
            await context.close();
          }
        }
    } finally {
      await browser.close();
    }
  }
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  if (error.stdout) report.commandOutput = String(error.stdout).slice(-4000);
  if (error.stderr) report.commandError = String(error.stderr).slice(-4000);
  save();
  throw error;
} finally {
  for (const server of servers) server.kill('SIGTERM');
  rmSync(consumer, { recursive: true, force: true });
}

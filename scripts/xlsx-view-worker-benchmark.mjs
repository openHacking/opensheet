import { chromium, firefox, webkit } from '@playwright/test';
import * as XLSX from 'xlsx';
import { createHash } from 'node:crypto';
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2],
  url = process.argv[3];
if (!output || !url || existsSync(output))
  throw new Error('Pass a new report path and production preview URL');
const sheet = XLSX.utils.aoa_to_sheet(Array.from({ length: 100000 }, (_, i) => [i]));
const original = XLSX.write(XLSX.utils.book_new(sheet, 'Data'), {
  type: 'array',
  bookType: 'xlsx',
});
const needle = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
const active =
  '<sheetViews><sheetView workbookViewId="0" tabSelected="true"><selection activeCell="D8" sqref="D8"/></sheetView></sheetViews>';
const make = (replacement) => {
  const archive = XLSX.CFB.read(new Uint8Array(original), { type: 'array' });
  const index = archive.FullPaths.findIndex(
    (path) => path === 'Root Entry/xl/worksheets/sheet1.xml',
  );
  if (index < 0) throw new Error('Missing worksheet');
  const xml = new TextDecoder('utf-8', { fatal: true }).decode(archive.FileIndex[index].content);
  if (xml.split(needle).length !== 2) throw new Error('Unexpected source worksheet view');
  XLSX.CFB.utils.cfb_add(
    archive,
    '/xl/worksheets/sheet1.xml',
    new TextEncoder().encode(xml.replace(needle, replacement)),
  );
  return Buffer.from(XLSX.CFB.write(archive, { type: 'array', fileType: 'zip' }));
};
const inputs = { control: make(needle), active: make(active) };
const archiveParts = (input) => {
  const archive = XLSX.CFB.read(new Uint8Array(input), { type: 'array' });
  return new Map(archive.FullPaths.map((path, index) => [path, archive.FileIndex[index].content]));
};
const controlParts = archiveParts(inputs.control),
  activeParts = archiveParts(inputs.active);
if (controlParts.size !== activeParts.size) throw new Error('Fixture archive entries differ');
for (const [path, control] of controlParts) {
  const candidate = activeParts.get(path);
  if (!candidate && !control) continue;
  if (!candidate || !control) throw new Error('Fixture archive entry missing');
  if (path === 'Root Entry/xl/worksheets/sheet1.xml') {
    const decoder = new TextDecoder('utf-8', { fatal: true });
    if (decoder.decode(candidate).replace(active, needle) !== decoder.decode(control))
      throw new Error('Worksheet differs beyond view');
  } else if (!Buffer.from(control).equals(Buffer.from(candidate)))
    throw new Error('Fixture archive differs beyond view');
}
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const assets = readdirSync('apps/playground/dist/assets').filter((name) =>
  /^file\.worker-.*\.js$/.test(name),
);
if (assets.length !== 1) throw new Error('Expected exactly one file Worker bundle');
const workerPath = `assets/${assets[0]}`,
  workerBytes = readFileSync(`apps/playground/dist/${workerPath}`);
const served = await fetch(new URL(workerPath, url));
if (!served.ok || hash(new Uint8Array(await served.arrayBuffer())) !== hash(workerBytes))
  throw new Error('Served file Worker differs from local build');
const report = {
  startedAt: new Date().toISOString(),
  url,
  source: sourceFingerprint(),
  localBuild: localBuildFingerprint(),
  worker: { path: workerPath, sha256: hash(workerBytes) },
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  inputs: Object.fromEntries(
    Object.entries(inputs).map(([name, bytes]) => [
      name,
      { bytes: bytes.length, sha256: hash(bytes) },
    ]),
  ),
  fixtureComparison:
    'Archive part contents match except for the expected worksheet view substitution',
  scope:
    'Production file Worker, synthetic XLSX with 100,000 numeric cells. Three alternating pairs per headless browser after warmup; active view differs from control and intentionally generates two warnings. Times include ZIP preflight, SheetJS parsing, adapter conversion and Worker response; every input value and expected report checked outside timing. This is not full UI import, persistence, peak memory, a controlled before/after build or reference-device acceptance.',
  results: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), pairs: [] };
  report.results.push(result);
  try {
    const page = await browser.newPage();
    for (const [name, bytes] of Object.entries(inputs))
      await page.route(`**/xlsx-view-bench-${name}.xlsx`, (route) =>
        route.fulfill({
          body: bytes,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
      );
    await page.goto(url);
    await page.waitForFunction(
      () =>
        !!window.opensheet?.selection.get() && !document.querySelector('#scene-select').disabled,
    );
    await page.evaluate(() => window.opensheet.getGrid().ready());
    const measurement = await page.evaluate(
      async ({ workerPath }) => {
        const fixtures = {};
        for (const name of ['control', 'active'])
          fixtures[name] = await (await fetch(`/xlsx-view-bench-${name}.xlsx`)).arrayBuffer();
        const worker = new Worker(`/${workerPath}`, { type: 'module' });
        const run = async (name) => {
          const buffer = fixtures[name].slice(0),
            start = performance.now();
          const data = await new Promise((resolve, reject) => {
            const cleanup = () => {
              clearTimeout(timeout);
              worker.removeEventListener('message', message);
              worker.removeEventListener('error', failure);
              worker.removeEventListener('messageerror', messageFailure);
            };
            const message = (event) => {
              cleanup();
              resolve(event.data);
            };
            const failure = (event) => {
              cleanup();
              reject(new Error(event.message));
            };
            const messageFailure = () => {
              cleanup();
              reject(new Error('Worker message deserialization failed'));
            };
            const timeout = setTimeout(() => {
              cleanup();
              reject(new Error('Worker benchmark timed out'));
            }, 30000);
            worker.addEventListener('message', message, { once: true });
            worker.addEventListener('error', failure, { once: true });
            worker.addEventListener('messageerror', messageFailure, { once: true });
            worker.postMessage({ buffer, name: 'benchmark.xlsx', profile: true }, [buffer]);
          });
          const roundTripMs = performance.now() - start;
          if (data.error || !data.file || !data.profile)
            throw new Error(data.error ?? 'Missing Worker result');
          const codes = data.report.issues.map((issue) => issue.code);
          for (const code of ['WORKSHEET_SELECTION_DROPPED', 'WORKSHEET_ACTIVE_TAB_DROPPED'])
            if (codes.includes(code) !== (name === 'active')) throw new Error('Wrong view report');
          const seen = new Uint8Array(100000);
          let count = 0;
          for (const block of data.file.sheets[0].blocks) {
            if (block.column !== 0 || !block.data.numbers) throw new Error('Unexpected page');
            const { positions, values } = block.data.numbers;
            if (positions.length !== values.length) throw new Error('Wrong column length');
            for (let i = 0; i < positions.length; i++) {
              const offset = positions[i];
              if (offset % 32 !== 0) throw new Error('Wrong cell column');
              const row = block.row * 64 + offset / 32;
              if (row >= seen.length || seen[row] || values[i] !== row)
                throw new Error('Wrong cell value');
              seen[row] = 1;
              count++;
            }
          }
          if (count !== seen.length) throw new Error('Missing cells');
          return { roundTripMs, ...data.profile, checkedCells: count };
        };
        try {
          await run('control');
          await run('active');
          const pairs = [];
          for (let i = 0; i < 3; i++) {
            const pair = {};
            for (const name of i % 2 ? ['active', 'control'] : ['control', 'active'])
              pair[name] = await run(name);
            pairs.push(pair);
          }
          return pairs;
        } finally {
          worker.terminate();
        }
      },
      { workerPath },
    );
    result.pairs = measurement;
    save();
  } catch (error) {
    result.error = String(error);
    throw error;
  } finally {
    await browser.close();
    save();
  }
}
if (JSON.stringify(localBuildFingerprint()) !== JSON.stringify(report.localBuild)) {
  report.error = 'Local build changed during measurement';
  save();
  throw new Error(report.error);
}
report.finishedAt = new Date().toISOString();
save();

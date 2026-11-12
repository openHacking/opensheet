import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createWorkbookFile, encodeFileCells } from '../packages/core/dist/index.js';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/node-consumer-acceptance.mjs NEW_REPORT.json');
const root = resolve('.');
const hash = (data) => createHash('sha256').update(data).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  source: sourceFingerprint(),
  scope:
    'Isolated temporary consumer installs locally packed formula and core tarballs using the offline pnpm store. It exercises pure JSON reads; it is not a supported full Node workbook runtime.',
  packages: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const consumer = mkdtempSync(join(tmpdir(), 'opensheet-node-consumer-'));
try {
  mkdirSync(join(consumer, 'tarballs'));
  const dependencies = {},
    overrides = {};
  for (const folder of ['formula', 'core']) {
    const directory = join(root, 'packages', folder);
    const manifest = JSON.parse(readFileSync(join(directory, 'package.json'), 'utf8'));
    const archive = join(consumer, 'tarballs', `${folder}.tgz`);
    execFileSync('rtk', ['proxy', 'pnpm', 'pack', '--out', archive], {
      cwd: directory,
      stdio: 'pipe',
    });
    const reference = `file:./tarballs/${folder}.tgz`;
    dependencies[manifest.name] = overrides[manifest.name] = reference;
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
        name: 'node-reader-consumer',
        private: true,
        type: 'module',
        dependencies,
        pnpm: { overrides },
      },
      null,
      2,
    ),
  );
  execFileSync(
    'rtk',
    ['proxy', 'pnpm', 'install', '--offline', '--ignore-scripts', '--strict-peer-dependencies'],
    { cwd: consumer, stdio: 'pipe', timeout: 120000 },
  );
  const lock = readFileSync(join(consumer, 'pnpm-lock.yaml'));
  const lockPath = output.replace(/\.json$/, '') + '.consumer-lock.yaml';
  writeFileSync(lockPath, lock, { flag: 'wx' });
  report.lock = { file: lockPath, sha256: hash(lock) };
  const example = join(consumer, 'node-file-reader.mjs');
  copyFileSync(join(root, 'examples/node-file-reader.mjs'), example);
  report.exampleSha256 = hash(readFileSync(example));
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 3, columns: 3 }] });
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        [
          { rowId: 0, columnId: 0, input: { type: 'number', value: 2 } },
          { rowId: 0, columnId: 1, input: { type: 'number', value: 3 } },
          { rowId: 0, columnId: 2, input: { type: 'formula', expression: 'SUM(A1:B1)' } },
        ],
        new Map(),
      ),
    },
  ];
  const fixture = join(consumer, 'book.json');
  writeFileSync(fixture, JSON.stringify(file));
  const read = execFileSync(process.execPath, [example, fixture, 'data', 'A1:C1'], {
    cwd: consumer,
    encoding: 'utf8',
  });
  assert.deepEqual(JSON.parse(read), { sheet: 'Data', range: 'A1:C1', values: [[2, 3, 5]] });
  for (const [name, range, expected] of [
    ['missing', 'A1', /Unknown sheet/],
    ['Data', 'A1:A10001', /at most 10,000 cells/],
  ]) {
    const result = spawnSync(process.execPath, [example, fixture, name, range], {
      cwd: consumer,
      encoding: 'utf8',
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, expected);
  }
  const runtime = JSON.parse(
    execFileSync(
      process.execPath,
      [
        '--input-type=module',
        '-e',
        `
    import { readFileSync } from 'node:fs';
    import { createWorkbook, FileReader } from '@opensheetjs/core';
    const globals = { document: typeof globalThis.document, Worker: typeof globalThis.Worker, indexedDB: typeof globalThis.indexedDB };
    const file = JSON.parse(readFileSync('book.json','utf8'));
    const reader = new FileReader(file);
    const batches = [...reader.getSheetById(file.sheets[0].id).range('A1:C3').batches({maxCells:3})];
    reader.dispose();
    let code;
    try { await createWorkbook(); code = 'unexpected success'; } catch (error) { code = error.code; }
    console.log(JSON.stringify({ globals, code, batches, resolved: import.meta.resolve('@opensheetjs/core') }));
  `,
      ],
      { cwd: consumer, encoding: 'utf8' },
    ),
  );
  assert.deepEqual(runtime.globals, {
    document: 'undefined',
    Worker: 'undefined',
    indexedDB: 'undefined',
  });
  assert.equal(runtime.code, 'WORKER_UNAVAILABLE');
  assert.deepEqual(
    runtime.batches.map((item) => item.values),
    [[[2, 3, 5]], [[null, null, null]], [[null, null, null]]],
  );
  const resolvedPath = realpathSync(fileURLToPath(runtime.resolved));
  const fromConsumer = relative(realpathSync(consumer), resolvedPath);
  assert.ok(fromConsumer && !fromConsumer.startsWith('..') && !isAbsolute(fromConsumer));
  report.results = {
    example: 'passed',
    invalidSheet: 'rejected',
    oversizedRange: 'rejected',
    batches: 'passed',
    runtime: { globals: runtime.globals, errorCode: runtime.code },
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report.results));
} catch (error) {
  report.error = String(error);
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  throw error;
} finally {
  rmSync(consumer, { recursive: true, force: true });
}

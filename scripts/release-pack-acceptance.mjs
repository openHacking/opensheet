import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/release-pack-acceptance.mjs NEW_REPORT.json');
const root = process.cwd(),
  consumer = mkdtempSync(join(tmpdir(), 'opensheet-release-pack-'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const version = JSON.parse(readFileSync('package.json', 'utf8')).version;
const report = {
  startedAt: new Date().toISOString(),
  version,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  node: process.version,
  packages: [],
  scope:
    'Two staged pnpm packs of each built workspace package, compared to direct pnpm pack, plus an isolated consumer installed with prefer-offline. Byte reproducibility is observed on this host, not independently built or signed provenance. No publication, browser lifecycle, cross-version migration or registry installation of OpenSheet.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
try {
  const directories = readdirSync(join(root, 'packages'), { withFileTypes: true })
    .filter((d) => d.isDirectory())
    .map((d) => d.name)
    .sort();
  mkdirSync(join(consumer, 'tarballs'));
  const dependencies = { react: '19.2.0', '@types/react': '19.2.2', vue: '3.5.42' },
    overrides = {};
  for (const folder of directories) {
    const packageDir = join(root, 'packages', folder);
    const original = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'));
    assert.equal(original.version, version, folder);
    const archive = join(consumer, 'tarballs', `${folder}.tgz`),
      repeat = join(consumer, 'tarballs', `${folder}-repeat.tgz`),
      direct = join(consumer, 'tarballs', `${folder}-direct.tgz`);
    execFileSync('pnpm', ['pack', '--out', direct], {
      cwd: packageDir,
      stdio: 'pipe',
    });
    for (const target of [archive, repeat])
      execFileSync(process.execPath, [join(root, 'scripts/pack-candidate.mjs'), folder, target], {
        cwd: root,
        stdio: 'pipe',
      });
    const first = readFileSync(archive),
      second = readFileSync(repeat);
    const files = execFileSync('tar', ['-tzf', archive], { encoding: 'utf8' })
      .trim()
      .split('\n')
      .sort();
    const manifest = JSON.parse(
      execFileSync('tar', ['-xOf', archive, 'package/package.json'], {
        encoding: 'utf8',
      }),
    );
    const directFiles = execFileSync('tar', ['-tzf', direct], { encoding: 'utf8' })
      .trim()
      .split('\n')
      .sort();
    assert.deepEqual(files, directFiles, `${folder}: staged file list differs`);
    const directManifest = JSON.parse(
      execFileSync('tar', ['-xOf', direct, 'package/package.json'], {
        encoding: 'utf8',
      }),
    );
    assert.deepEqual(manifest, directManifest, `${folder}: staged manifest differs`);
    for (const file of files) {
      if (file === 'package/package.json' || file.endsWith('/')) continue;
      const content = (path) =>
        execFileSync('tar', ['-xOf', path, file], { maxBuffer: 20 * 1024 * 1024 });
      assert.ok(
        content(archive).equals(content(direct)),
        `${folder}: staged file differs: ${file}`,
      );
    }
    assert.equal(manifest.name, original.name);
    assert.equal(manifest.version, version);
    for (const required of [
      'package/package.json',
      'package/README.md',
      'package/LICENSE',
      'package/dist/index.js',
      'package/dist/index.d.ts',
    ])
      assert.ok(files.includes(required), `${folder}: missing ${required}`);
    if (folder === 'core')
      assert.ok(files.includes('package/dist/engine.worker.js'), 'missing engine Worker');
    if (folder === 'opensheet')
      assert.ok(files.includes('package/dist/style.css'), 'missing editor CSS');
    for (const [name, spec] of Object.entries(manifest.dependencies ?? {})) {
      assert.ok(!spec.startsWith('workspace:'), `${folder}: unconverted ${name}`);
      if (name === 'opensheet' || name.startsWith('@opensheetjs/'))
        assert.equal(spec, version, `${folder}: ${name}`);
    }
    const reference = `file:./tarballs/${folder}.tgz`;
    dependencies[manifest.name] = overrides[manifest.name] = reference;
    const item = {
      name: manifest.name,
      version: manifest.version,
      sha256: hash(first),
      repeatSha256: hash(second),
      identical: first.equals(second),
      files: files.length,
      bytes: first.length,
    };
    report.packages.push(item);
    save();
    assert.ok(item.identical, `${folder}: repeated archives differ`);
    console.log(`${manifest.name}: packed and inspected`);
  }
  writeFileSync(
    join(consumer, 'package.json'),
    JSON.stringify(
      {
        name: 'opensheet-release-consumer',
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
    'pnpm',
    ['install', '--prefer-offline', '--ignore-scripts', '--strict-peer-dependencies'],
    { cwd: consumer, stdio: 'pipe', encoding: 'utf8', timeout: 120000 },
  );
  const lock = readFileSync(join(consumer, 'pnpm-lock.yaml')),
    lockPath = output.replace(/\.json$/, '') + '.consumer-lock.yaml';
  writeFileSync(lockPath, lock, { flag: 'wx' });
  report.lock = { path: lockPath, sha256: hash(lock) };
  const names = report.packages.map((p) => p.name);
  const script = `
const names=${JSON.stringify(names)};
if(typeof document!=='undefined'||typeof Worker!=='undefined'||typeof indexedDB!=='undefined')throw Error('Unexpected browser globals');
for(const name of names) { const module=await import(name); if(!Object.keys(module).length)throw Error('Empty package '+name); }
const {createWorkbookFile,encodeFileCells,FileReader}=await import('@opensheetjs/core');
const {toSheetJS,fromSheetJS}=await import('@opensheetjs/adapter-sheetjs');
const {exportRange}=await import('@opensheetjs/formats');
const file=createWorkbookFile({sheets:[{name:'Data',rows:2,columns:2}]});
file.sheets[0].blocks=[{row:0,column:0,data:encodeFileCells([{rowId:0,columnId:0,input:{type:'number',value:7}},{rowId:0,columnId:1,input:{type:'formula',expression:'A1+1'}}],new Map())}];
const reader=new FileReader(file);const values=reader.getSheetById(file.sheets[0].id).range('A1:B1').getValues();reader.dispose();
if(JSON.stringify(values)!=='[[7,8]]')throw Error('Bad formula result');
const converted=fromSheetJS(toSheetJS(file).workbook).file;if(converted.sheets.length!==1)throw Error('Bad adapter conversion');
if(!exportRange(file,{sheetId:file.sheets[0].id,format:'csv'}).text.includes('7'))throw Error('Bad table export');
console.log(JSON.stringify({imported:names.length,values}));`;
  writeFileSync(join(consumer, 'check.mjs'), script);
  const result = JSON.parse(
    execFileSync(process.execPath, [join(consumer, 'check.mjs')], {
      cwd: consumer,
      encoding: 'utf8',
    }),
  );
  assert.deepEqual(result, { imported: directories.length, values: [[7, 8]] });
  const types = `
import {createWorkbookFile, FileReader, type WorkbookFile} from '@opensheetjs/core';
import {parseFormula, evaluateFormula} from '@opensheetjs/formula';
import {toSheetJS} from '@opensheetjs/adapter-sheetjs';
import {exportRange} from '@opensheetjs/formats';
import {createOpenSheet} from 'opensheet';
import {OpenSheetView as ReactView} from '@opensheetjs/react';
import {OpenSheetView as VueView} from '@opensheetjs/vue';
import {CanvasGrid} from '@opensheetjs/renderer';
import {definePlugin} from '@opensheetjs/plugin-sdk';
import {normalizedSnapshot} from '@opensheetjs/testing';
const file: WorkbookFile = createWorkbookFile();
const reader = new FileReader(file);
void [reader,toSheetJS(file),exportRange,createOpenSheet,ReactView,VueView,CanvasGrid,definePlugin,normalizedSnapshot(file),evaluateFormula(parseFormula('1+1'),()=>null)];
`;
  writeFileSync(join(consumer, 'consumer.mts'), types);
  execFileSync(
    process.execPath,
    [
      join(root, 'node_modules/typescript/bin/tsc'),
      '--noEmit',
      '--module',
      'NodeNext',
      '--moduleResolution',
      'NodeNext',
      '--target',
      'ES2022',
      '--strict',
      'consumer.mts',
    ],
    { cwd: consumer, encoding: 'utf8', stdio: 'pipe', timeout: 120000 },
  );
  report.consumer = { ...result, types: 'passed' };
  report.finishedAt = new Date().toISOString();
  save();
  console.log('Isolated consumer: passed');
} catch (error) {
  report.error = String(error);
  save();
  throw error;
} finally {
  rmSync(consumer, { recursive: true, force: true });
}

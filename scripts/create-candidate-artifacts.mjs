import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const [destinationArg, cleanReportArg, consumerReportArg] = process.argv.slice(2);
if (!destinationArg || !cleanReportArg || !consumerReportArg)
  throw new Error(
    'Usage: node scripts/create-candidate-artifacts.mjs NEW_DIRECTORY CLEAN_REPORT CONSUMER_REPORT',
  );
const destination = resolve(destinationArg);
const root = process.cwd();
const cleanBytes = readFileSync(cleanReportArg);
const consumerBytes = readFileSync(consumerReportArg);
const clean = JSON.parse(cleanBytes),
  consumer = JSON.parse(consumerBytes);
const version = JSON.parse(readFileSync('package.json', 'utf8')).version;
assert.equal(clean.identical, true, 'Independent-directory build did not pass');
assert.equal(consumer.version, version, 'Consumer report version differs');
assert.equal(consumer.consumer?.types, 'passed', 'Consumer type check did not pass');
const directories = readdirSync('packages', { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort();
assert.equal(clean.runs?.length, 2, 'Expected two build runs');
assert.equal(consumer.packages?.length, directories.length, 'Consumer package count differs');
for (const run of clean.runs)
  assert.equal(run.packages?.length, directories.length, 'Build package count differs');
const sha256 = (data) => createHash('sha256').update(data).digest('hex');
const manifest = {
  version,
  createdAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  evidence: {
    clean: { path: cleanReportArg, sha256: sha256(cleanBytes) },
    consumer: { path: consumerReportArg, sha256: sha256(consumerBytes) },
  },
  scope:
    'Unpublished development archives verified against same-host independent-directory builds and installed-consumer results; no signing, registry publication, cross-platform or release approval.',
  packages: [],
};
// mkdir without recursive is intentional: a failure never removes user-owned content.
mkdirSync(destination);
try {
  for (const folder of directories) {
    const packageManifest = JSON.parse(
      readFileSync(join('packages', folder, 'package.json'), 'utf8'),
    );
    assert.equal(packageManifest.version, version, `${folder}: version mismatch`);
    const file = `${folder}.tgz`,
      output = join(destination, file);
    execFileSync(process.execPath, [join(root, 'scripts/pack-candidate.mjs'), folder, output], {
      cwd: root,
      encoding: 'utf8',
      stdio: 'pipe',
      timeout: 120000,
    });
    const bytes = readFileSync(output),
      hash = sha256(bytes);
    const inspected = consumer.packages.find((item) => item.name === packageManifest.name);
    assert.ok(inspected?.identical, `${folder}: consumer archive was not reproducible`);
    assert.equal(hash, inspected.sha256, `${folder}: consumer-verified archive differs`);
    for (const run of clean.runs)
      assert.equal(
        hash,
        run.packages.find((item) => item.name === folder)?.sha256,
        `${folder}: independent build differs`,
      );
    manifest.packages.push({
      name: packageManifest.name,
      version,
      file,
      bytes: bytes.length,
      sha256: hash,
    });
    console.log(`${packageManifest.name}: verified`);
  }
  writeFileSync(join(destination, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n', {
    flag: 'wx',
  });
  writeFileSync(
    join(destination, 'SHA256SUMS'),
    manifest.packages.map((item) => `${item.sha256}  ${item.file}`).join('\n') + '\n',
    { flag: 'wx' },
  );
} catch (error) {
  rmSync(destination, { recursive: true, force: true });
  throw error;
}

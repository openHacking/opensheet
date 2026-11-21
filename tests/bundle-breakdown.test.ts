import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';

it('retains reproducible bundle measurements and the lazy XLSX boundary', () => {
  const names = ['final', 'repeat', 'rebuilt', 'served'];
  const reports = names.map((name) =>
    JSON.parse(readFileSync(`benchmarks/bundle-breakdown-${name}-2026-11-22.json`, 'utf8')),
  );
  for (const report of reports) {
    expect(report.error).toBeUndefined();
    expect(report.lock.frozenReinstall).toBe('passed');
    expect(report.lazyIsolation.xlsxOutputs.length).toBeGreaterThan(0);
    expect(createHash('sha256').update(readFileSync(report.lock.path)).digest('hex')).toBe(
      report.lock.sha256,
    );
    expect(report.packages).toHaveLength(6);
    for (const scenario of Object.values(report.scenarios) as Record<string, any>[]) {
      for (const output of Object.values(scenario) as any[]) {
        expect(output.bytes).toBeGreaterThan(0);
        expect(output.gzipBytes).toBeGreaterThan(0);
        expect(output.sha256).toMatch(/^[a-f0-9]{64}$/);
        const accounted = (Object.values(output.contributions) as number[]).reduce(
          (a, b) => a + b,
          0,
        );
        expect(accounted).toBeLessThanOrEqual(output.bytes);
      }
    }
    const outputs = report.scenarios.lazy;
    const entry = Object.keys(outputs).find((name) => outputs[name].entryPoint === 'lazy.js');
    expect(entry).toBeDefined();
    const initial = new Set<string>();
    const visit = (file: string) => {
      if (initial.has(file)) return;
      initial.add(file);
      for (const dependency of outputs[file].imports)
        if (!dependency.external && dependency.kind !== 'dynamic-import') visit(dependency.path);
      // CSS is an emitted initial asset; it has no JS imports in this fixture.
    };
    visit(entry!);
    const xlsx = Object.keys(outputs).filter((file) => outputs[file].contributions.xlsx > 0);
    expect(xlsx.length).toBeGreaterThan(0);
    for (const file of xlsx) expect(initial.has(file)).toBe(false);
    expect(report.lazyIsolation.status).toBe('passed');
  }
  const fingerprints = (report: any) =>
    Object.fromEntries(
      Object.entries(report.scenarios).map(([name, outputs]) => [
        name,
        Object.fromEntries(
          Object.entries(outputs as Record<string, any>).map(([file, info]) => [file, info.sha256]),
        ),
      ]),
    );
  expect(fingerprints(reports[1])).toEqual(fingerprints(reports[0]));
  expect(fingerprints(reports[2])).toEqual(fingerprints(reports[0]));
  expect(fingerprints(reports[3])).toEqual(fingerprints(reports[0]));
  const served = reports[3];
  expect(served.runtime.workerSha256).toBe(
    served.scenarios.worker['out/worker/engine.worker.js'].sha256,
  );
  expect(served.runtime.results.map((result: any) => result.engine)).toEqual([
    'chromium',
    'firefox',
    'webkit',
  ]);
  const xlsxPaths = served.lazyIsolation.xlsxOutputs.map(
    (file: string) => served.runtime.prefix + file.split('/').at(-1),
  );
  for (const result of served.runtime.results) {
    expect(result.before).toEqual({ values: [[6, 7, 13]], height: 500, stylesheets: 1 });
    expect(result.beforeRequests).toContain(`${served.runtime.prefix}engine.worker.js`);
    for (const path of xlsxPaths) {
      expect(result.beforeRequests).not.toContain(path);
      expect(result.afterRequests.filter((request: string) => request === path)).toHaveLength(1);
    }
    expect(result.converted).toMatchObject({
      values: [[6, 7, 13]],
      reopened: [[6, 7, 13]],
      readOnly: false,
    });
    expect(result.converted.bytes).toBeGreaterThan(0);
    expect(result.errors).toEqual([]);
    expect(result.failures).toEqual([]);
  }
});

it('records enforced CSP controls and lazy XLSX conversion in three browsers', () => {
  const strict = JSON.parse(
    readFileSync('benchmarks/bundle-breakdown-csp-verified-2026-11-22.json', 'utf8'),
  );
  const regular = JSON.parse(
    readFileSync('benchmarks/bundle-breakdown-nocsp-verified-2026-11-22.json', 'utf8'),
  );
  expect(strict.strictCsp).toBe(true);
  expect(regular.strictCsp).toBe(false);
  expect(strict.cspShimSha256).toBe(
    createHash('sha256').update(readFileSync('examples/framework-host/csp-zod.js')).digest('hex'),
  );
  expect(strict.runtime.csp).toContain("script-src 'self'");
  expect(strict.runtime.csp).toContain("worker-src 'self'");
  expect(strict.runtime.csp).not.toContain('unsafe-eval');
  expect(strict.runtime.csp).not.toContain('unsafe-inline');
  expect(regular.runtime.csp).toBeNull();
  for (const report of [strict, regular]) {
    expect(report.lock.frozenReinstall).toBe('passed');
    expect(report.lazyIsolation.xlsxOutputs.length).toBeGreaterThan(0);
    expect(createHash('sha256').update(readFileSync(report.lock.path)).digest('hex')).toBe(
      report.lock.sha256,
    );
    expect(report.runtime.workerSha256).toBe(
      report.scenarios.worker['out/worker/engine.worker.js'].sha256,
    );
    expect(report.runtime.results.map((result: any) => result.engine)).toEqual([
      'chromium',
      'firefox',
      'webkit',
    ]);
    for (const result of report.runtime.results) {
      expect(result.converted).toMatchObject({
        values: [[6, 7, 13]],
        reopened: [[6, 7, 13]],
        readOnly: false,
      });
      expect(result.before.stylesheets).toBe(1);
      expect(result.errors).toEqual([]);
      expect(result.failures).toEqual([]);
      for (const file of report.lazyIsolation.xlsxOutputs) {
        const path = report.runtime.prefix + file.split('/').at(-1);
        expect(result.beforeRequests).not.toContain(path);
        expect(result.afterRequests.filter((request: string) => request === path)).toHaveLength(1);
      }
      if (report.strictCsp) {
        expect(result.cspControls).toEqual({ pageBlocked: true, workerBlocked: true });
        for (const event of result.cspEvents) {
          expect(event.blocked).toBe('eval');
          expect(event.source).toContain('/probe.js');
        }
      } else expect(result.cspControls).toBeNull();
    }
  }
});

it('records Worker CSP telemetry with an engine-level negative control', () => {
  const observed = JSON.parse(
    readFileSync('benchmarks/bundle-breakdown-worker-observed-2026-11-22.json', 'utf8'),
  );
  const probe = JSON.parse(
    readFileSync('benchmarks/bundle-breakdown-worker-probe-2026-11-22.json', 'utf8'),
  );
  const regular = JSON.parse(
    readFileSync('benchmarks/bundle-breakdown-worker-nocsp-2026-11-22.json', 'utf8'),
  );
  const source = readFileSync('scripts/bundle-runtime.mjs', 'utf8');
  const telemetry = source.match(/export const workerTelemetry = `([^`]+)`/)?.[1];
  expect(telemetry).toBeDefined();
  const telemetryHash = createHash('sha256').update(telemetry!).digest('hex');
  expect(observed.engineCspProbe).toBe(false);
  expect(probe.engineCspProbe).toBe(true);
  expect(regular.strictCsp).toBe(false);
  expect(observed.runtime.workerTelemetrySha256).toBe(telemetryHash);
  expect(probe.runtime.workerTelemetrySha256).toBe(telemetryHash);
  expect(observed.workerBannerSha256).not.toBe(probe.workerBannerSha256);
  expect(observed.runtime.workerSha256).not.toBe(probe.runtime.workerSha256);
  for (const report of [observed, probe, regular]) {
    expect(report.error).toBeUndefined();
    expect(report.lock.frozenReinstall).toBe('passed');
    expect(createHash('sha256').update(readFileSync(report.lock.path)).digest('hex')).toBe(
      report.lock.sha256,
    );
    expect(report.runtime.workerSha256).toBe(
      report.scenarios.worker['out/worker/engine.worker.js'].sha256,
    );
    expect(report.runtime.results.map((result: any) => result.engine)).toEqual([
      'chromium',
      'firefox',
      'webkit',
    ]);
    for (const result of report.runtime.results) {
      expect(result.converted).toMatchObject({
        values: [[6, 7, 13]],
        reopened: [[6, 7, 13]],
        readOnly: false,
      });
      expect(result.errors).toEqual([]);
      expect(result.failures).toEqual([]);
      if (report.strictCsp)
        expect(result.cspControls).toEqual({ pageBlocked: true, workerBlocked: true });
    }
  }
  for (const report of [observed, probe]) {
    for (const result of report.runtime.results) {
      const negative = result.workerCspEvents.filter((event: any) =>
        event.path.endsWith('/probe.worker.js'),
      );
      if (result.engine === 'webkit') expect(negative).toEqual([]);
      else
        expect(negative).toEqual(
          expect.arrayContaining([
            expect.objectContaining({ blocked: 'eval', directive: 'script-src' }),
          ]),
        );
    }
  }
  for (const result of observed.runtime.results)
    expect(
      result.workerCspEvents.filter((event: any) => event.path.endsWith('/engine.worker.js')),
    ).toEqual([]);
  for (const result of probe.runtime.results) {
    const engineEvents = result.workerCspEvents.filter((event: any) =>
      event.path.endsWith('/engine.worker.js'),
    );
    if (result.engine === 'webkit') expect(engineEvents).toEqual([]);
    else
      expect(engineEvents).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ blocked: 'eval', directive: 'script-src' }),
        ]),
      );
  }
});

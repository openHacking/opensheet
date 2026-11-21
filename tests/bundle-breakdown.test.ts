import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';

it('retains reproducible bundle measurements and the lazy XLSX boundary', () => {
  const names = ['final', 'repeat', 'rebuilt'];
  const reports = names.map((name) =>
    JSON.parse(readFileSync(`benchmarks/bundle-breakdown-${name}-2026-11-22.json`, 'utf8')),
  );
  for (const report of reports) {
    expect(report.error).toBeUndefined();
    expect(report.lock.frozenReinstall).toBe('passed');
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
});

import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';

it('pins independent 1900 and 1904 date round trips and recalculated formula caches', () => {
  const report = JSON.parse(
    readFileSync('benchmarks/date-system-reader-verified-2026-11-22.json', 'utf8'),
  );
  expect(report.error).toBeUndefined();
  expect(report.readerVersion).toContain('LibreOfficeDev 26.8');
  expect(report.systems.map((system: { system: string }) => system.system)).toEqual([
    '1900',
    '1904',
  ]);
  for (const system of report.systems) {
    expect(system.files).toHaveLength(4);
    for (const file of system.files)
      expect(createHash('sha256').update(readFileSync(file.path)).digest('hex')).toBe(file.sha256);
    expect(system.independent.reader).toBe('openpyxl 3.1.5');
    expect(system.independent.files).toHaveLength(4);
    for (const [index, observation] of system.independent.files.entries()) {
      expect(observation.system).toBe(system.system);
      expect(observation.dates).toEqual([
        '1904-01-02T00:00:00',
        '2024-02-29T12:30:00',
        '2100-03-01T00:00:00',
      ]);
      expect(observation.formula).toBe('=A2-A1');
      if (index < 2) expect(observation.cached).toBeNull();
      else expect(observation.cached).toBeCloseTo(43888.5208333333, 8);
    }
    const [source, exported, control, reopened] = system.independent.files;
    expect(exported.rawSerials).toEqual(source.rawSerials);
    expect(control.rawSerials).toEqual(reopened.rawSerials);
    expect(control.rawSerials[1]).not.toBe(source.rawSerials[1]);
    expect(reopened.formats).toEqual([
      'yyyy\\-mm\\-dd',
      'yyyy\\-mm\\-dd\\ hh:mm',
      'yyyy\\-mm\\-dd',
    ]);
    expect(system.exportReport.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'RECALC_REQUIRED', action: 'approximated' }),
        expect.objectContaining({ code: 'UNSUPPORTED_METADATA', action: 'dropped' }),
      ]),
    );
    for (const observation of system.openSheet) {
      expect(observation.system).toBe(system.system);
      expect(observation.display).toEqual([['1904-01-02'], ['2024-02-29 12:30'], ['2100-03-01']]);
      expect(observation.formula).toEqual({ type: 'formula', expression: 'A2-A1' });
      expect(observation.calculated).toBeCloseTo(43888.5208333333, 8);
    }
  }
});

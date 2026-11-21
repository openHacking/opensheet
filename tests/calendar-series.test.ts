import * as XLSX from 'xlsx';
import { expect, it } from 'vitest';
import { calendarSeries } from '../packages/core/src/calendar-series.js';

it('anchors month and year steps to the original day, including the 1900 compatibility day', () => {
  // 1900-01-31 -> 1900-02-29 -> 1900-03-31, then back across the same boundary.
  expect(calendarSeries(31, 1, 3, 'month', '1900')).toEqual([31, 60, 91]);
  expect(calendarSeries(91, -1, 3, 'month', '1900')).toEqual([91, 60, 31]);
  // Leap day clamps in common years and returns to leap day in the next leap year.
  expect(calendarSeries(43890, 1, 5, 'year', '1900')).toEqual([43890, 44255, 44620, 44985, 45351]);
  expect(calendarSeries(60, 1, 5, 'year', '1900')).toEqual([60, 425, 790, 1155, 1521]);
});

it('supports 1904 dates and rejects invalid or out-of-range sequences before generation', () => {
  expect(calendarSeries(30, 1, 3, 'month', '1904')).toEqual([30, 59, 90]);
  expect(calendarSeries(30, 0, 3, 'year', '1904')).toEqual([30, 30, 30]);
  for (const run of [
    () => calendarSeries(0, 1, 2, 'month', '1900'),
    () => calendarSeries(30, 0.5, 2, 'month', '1904'),
    () => calendarSeries(0, -1, 2, 'month', '1904'),
    () => calendarSeries(2958465, 1, 2, 'year', '1900'),
    () => calendarSeries(30, Number.MAX_SAFE_INTEGER, 2, 'year', '1904'),
  ])
    expect(run).toThrow();
});

it('matches independently recalculated anchored calendar sequences', async () => {
  const { createHash } = await import('node:crypto');
  const { readFileSync } = await import('node:fs');
  const root = new URL('../benchmarks/', import.meta.url);
  const evidence = JSON.parse(
    readFileSync(new URL('calendar-series-reader-2026-11-21.json', root), 'utf8'),
  );
  for (const artifact of Object.values(evidence.artifacts) as { file: string; sha256: string }[])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  expect(evidence.cases).toHaveLength(14);
  expect(
    createHash('sha256')
      .update(readFileSync(new URL('../scripts/calendar-series-reader.py', import.meta.url)))
      .digest('hex'),
  ).toBe(evidence.scriptSha256);
  for (const system of ['1900', '1904'] as const) {
    const source = XLSX.read(
      readFileSync(new URL(evidence.artifacts[`${system}_source`].file, root)),
      { type: 'buffer' },
    );
    const saved = XLSX.read(
      readFileSync(new URL(evidence.artifacts[`${system}_saved`].file, root)),
      { type: 'buffer' },
    );
    expect(!!saved.Workbook?.WBProps?.date1904).toBe(system === '1904');
    for (const item of evidence.cases.filter(
      (entry: { system: string }) => entry.system === system,
    )) {
      expect(source.Sheets[item.case].B1.v).toBe(item.start);
      expect(item.values).toEqual(
        item.values.map((_: number, index: number) => saved.Sheets[item.case][`B${index + 4}`].v),
      );
      for (let index = 0; index < item.values.length; index++) {
        expect(`=${saved.Sheets[item.case][`B${index + 4}`].f}`).toBe(item.readerFormulas[index]);
        expect(item.readerFormulas[index]).toMatch(/^=EDATE\(/i);
      }
      expect(
        calendarSeries(item.start, item.step, item.values.length, item.unit, system),
        `${system}/${item.case}`,
      ).toEqual(item.values);
    }
  }
});

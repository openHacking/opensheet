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

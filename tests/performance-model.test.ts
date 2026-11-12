import { expect, it } from 'vitest';
import {
  capacityEditRow,
  capacityHorizontalScrollOffset,
  capacityScrollOffset,
} from '../apps/playground/src/performance-model.js';

it.each([100, 10000, 100000])('spreads measured edits across all %s rows', (rows) => {
  const positions = Array.from({ length: 100 }, (_, index) => capacityEditRow(index, rows));
  expect(positions[0]).toBe(0);
  expect(positions.at(-1)).toBe(rows - 1);
  expect(new Set(positions).size).toBe(100);
  for (let i = 1; i < positions.length; i++) expect(positions[i]).toBeGreaterThan(positions[i - 1]);
  for (let i = -5; i < 0; i++) expect(capacityEditRow(i, rows)).toBe(0);
});

it.each([0, 99, 12_000])('covers both horizontal edges over %s pixels', (max) => {
  const positions = Array.from({ length: 300 }, (_, index) =>
    capacityHorizontalScrollOffset(index, max),
  );
  expect(Math.min(...positions)).toBe(0);
  expect(Math.max(...positions)).toBe(max);
  expect(new Set(positions).size).toBe(Math.min(300, max + 1));
  for (let i = -30; i < 0; i++) expect(capacityHorizontalScrollOffset(i, max)).toBe(0);
});

it.each([0, 99, 2_400_000])('spreads measured scrolls across %s pixels', (max) => {
  const positions = Array.from({ length: 300 }, (_, index) => capacityScrollOffset(index, max));
  expect(positions[0]).toBe(0);
  expect(positions.at(-1)).toBe(max);
  for (let i = 1; i < positions.length; i++)
    expect(positions[i]).toBeGreaterThanOrEqual(positions[i - 1]);
  for (let i = -30; i < 0; i++) expect(capacityScrollOffset(i, max)).toBe(0);
});

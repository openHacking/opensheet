import { expect, it } from 'vitest';
import {
  definedNameIndex,
  referencesDefinedName,
} from '../packages/adapter-sheetjs/src/defined-name-scan.js';

it('detects unqualified names in the appropriate sheet scope without treating literals or references as names', () => {
  const index = definedNameIndex(
    [
      { Name: 'GlobalValue' },
      { Name: 'LocalValue', Sheet: 1 },
      { Name: 'A1' },
      { Name: 'TRUE' },
      { Name: 'A' },
      { Name: 'NAME' },
    ],
    2,
  );
  const matches = (expression: string, sheet = 0) =>
    referencesDefinedName(expression, index.global, index.local.get(sheet));
  for (const expression of ['GlobalValue', 'SUM(2,globalvalue)', '1+GlobalValue*2'])
    expect(matches(expression), expression).toBe(true);
  expect(matches('LocalValue', 1)).toBe(true);
  expect(matches('LocalValue', 0)).toBe(false);
  for (const expression of [
    '"GlobalValue"',
    'GlobalValue(1)',
    'GlobalValue!A1',
    "'Other'!GlobalValue",
    'Table[GlobalValue]',
    'A1',
    'TRUE',
    'OtherGlobalValue',
    '中GlobalValue',
    'GlobalValue中',
    'A$1',
    '#NAME?',
  ])
    expect(matches(expression), expression).toBe(false);
});

it('detects local and global names in known qualified sheets, excluding external and cell references', () => {
  const index = definedNameIndex(
    [{ Name: 'LocalValue', Sheet: 1 }, { Name: 'A', Sheet: 1 }, { Name: 'Shared' }],
    2,
  );
  const qualified = new Map([
    ["O'BRIEN", index.local.get(1)!],
    ['OTHER', index.local.get(1)!],
  ]);
  const matches = (expression: string) =>
    referencesDefinedName(expression, index.global, undefined, qualified);
  for (const expression of [
    "'O''Brien'!LocalValue",
    'Other!LocalValue+1',
    'SUM(other!localvalue,1)',
    'Other!Shared',
  ])
    expect(matches(expression), expression).toBe(true);
  for (const expression of [
    'Other!A1',
    'Other!A$1',
    'Other!$A$1',
    'Missing!Shared',
    'Missing!LocalValue',
    "'[book.xlsx]Other'!LocalValue",
    '[book.xlsx]Other!LocalValue',
    'Other!LocalValue(1)',
    '"Other!LocalValue"',
  ])
    expect(matches(expression), expression).toBe(false);
});

it('bounds the name index and discloses skipped or unrepresentable names', () => {
  const names = Array.from({ length: 8193 }, (_, i) => ({ Name: `name_${i}` }));
  const index = definedNameIndex(names, 1);
  expect(index.global.size).toBe(8192);
  expect(index.incomplete).toBe(true);
  expect(definedNameIndex([{ Name: '本地' }, { Name: 'Local', Sheet: 2 }], 1).incomplete).toBe(
    true,
  );
});

it('detects Unicode local and global names without matching fragments or quoted text', () => {
  const index = definedNameIndex(
    [{ Name: '预算' }, { Name: '总计', Sheet: 1 }, { Name: '𝒳值' }],
    2,
  );
  expect(index.incomplete).toBe(false);
  const qualified = new Map([
    ['数据', index.local.get(1)!],
    ['OTHER', new Set<string>()],
  ]);
  const matches = (value: string, sheet = 0) =>
    referencesDefinedName(value, index.global, index.local.get(sheet), qualified);
  for (const value of ['预算+1', 'SUM(预算,1)', '𝒳值+1', '数据!总计', "'数据'!预算", 'Other!预算'])
    expect(matches(value), value).toBe(true);
  expect(matches('总计', 1)).toBe(true);
  for (const value of [
    '总计',
    '预算表',
    '前预算',
    '"预算"',
    '预算(1)',
    'Table[预算]',
    "'[book.xlsx]数据'!预算",
    'Missing!预算',
  ])
    expect(matches(value), value).toBe(false);
});

it('detects backslash-prefixed names without treating substrings as references', () => {
  const index = definedNameIndex([{ Name: '\\Budget' }, { Name: '\\本地', Sheet: 1 }], 2);
  expect(index.incomplete).toBe(false);
  const qualified = new Map([['DATA', index.local.get(1)!]]);
  const matches = (value: string, sheet = 0) =>
    referencesDefinedName(value, index.global, index.local.get(sheet), qualified);
  for (const value of ['\\Budget+1', 'SUM(\\Budget,1)', 'Data!\\Budget', "'Data'!\\本地"])
    expect(matches(value), value).toBe(true);
  expect(matches('\\本地', 1)).toBe(true);
  for (const value of [
    'Budget',
    '"\\Budget"',
    'Prefix\\Budget',
    '\\BudgetExtra',
    '\\Budget(1)',
    'Table[\\Budget]',
    "'[other.xlsx]Data'!\\Budget",
  ])
    expect(matches(value), value).toBe(false);
});

it('handles long unterminated literals, brackets and error prefixes without a false match', () => {
  const names = new Set(['GLOBALVALUE']);
  for (const expression of [
    '"' + 'GlobalValue'.repeat(2500),
    "'" + 'GlobalValue'.repeat(2500),
    '['.repeat(32760) + 'GlobalValue',
    '#'.repeat(32760) + 'x',
  ])
    expect(referencesDefinedName(expression, names), expression.slice(0, 30)).toBe(false);
});

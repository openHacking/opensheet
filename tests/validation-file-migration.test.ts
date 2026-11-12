import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { encodeWorkbookFile, validateWorkbookFile } from '../packages/core/src/file-codec.js';
import {
  downgradeValidationFile,
  stageValidationDowngrade,
  stageValidationFile,
  stageValidationUpgrade,
  upgradeValidationFile,
} from '../packages/core/src/validation-file-migration.js';

it('stages a rule-bearing v4 file and preserves detached rule metadata', async () => {
  const source = upgradeValidationFile(fixture());
  source.sheets[0].validation.push({
    id: 'positive',
    range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 1 },
    rule: { kind: 'formula', expression: 'A1>0', allowBlank: false, policy: 'reject' },
  });
  source.sheets[1].validation.push({
    id: 'list',
    range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    rule: { kind: 'list', values: ['a', 'b'], allowBlank: true, policy: 'warn' },
  });
  const pages: unknown[] = [];
  const result = await stageValidationFile(source, async (_sheet, page) => {
    pages.push(page);
  });
  expect(pages).toEqual(source.sheets[0].blocks);
  expect(result.schemaVersion).toBe(4);
  expect(result.sheets.map((sheet) => sheet.validation)).toEqual(
    source.sheets.map((sheet) => sheet.validation),
  );
  expect(result.sheets.every((sheet) => !('blocks' in sheet))).toBe(true);
  source.sheets[0].validation[0].range.startRow = 3;
  expect(result.sheets[0].validation[0].range.startRow).toBe(0);
});

it('validates every rule before staging and withholds metadata on late page errors', async () => {
  const source = upgradeValidationFile(fixture());
  source.sheets[1].validation.push({
    id: 'bad',
    range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
    rule: { kind: 'formula', expression: 'A1+', allowBlank: false, policy: 'reject' },
  });
  let pages = 0;
  await expect(
    stageValidationFile(source, async () => {
      pages++;
    }),
  ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  expect(pages).toBe(0);
  source.sheets[1].validation = [];
  source.sheets[0].blocks.push(source.sheets[0].blocks[0]);
  await expect(
    stageValidationFile(source, async () => {
      pages++;
    }),
  ).rejects.toThrow('Duplicate page');
  expect(pages).toBe(2);
});

it('withholds rule-bearing metadata when staging is cancelled or its sink fails', async () => {
  const source = upgradeValidationFile(fixture());
  source.sheets[0].validation.push({
    id: 'required',
    range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    rule: { kind: 'number', operator: 'gte', first: 0, allowBlank: false, policy: 'reject' },
  });
  const before = structuredClone(source);
  const abort = new AbortController();
  let pages = 0;
  await expect(
    stageValidationFile(
      source,
      async () => {
        pages++;
        abort.abort();
      },
      abort.signal,
    ),
  ).rejects.toMatchObject({ name: 'AbortError' });
  expect(pages).toBe(1);
  await expect(
    stageValidationFile(source, async () => {
      throw new Error('staging failed');
    }),
  ).rejects.toThrow('staging failed');
  expect(source).toEqual(before);
});

it('passes cancellation to an in-flight page sink and waits for it to stop', async () => {
  const source = upgradeValidationFile(fixture());
  const controller = new AbortController();
  let started!: () => void;
  const entered = new Promise<void>((resolve) => {
    started = resolve;
  });
  let calls = 0,
    stopped = false;
  const operation = stageValidationFile(
    source,
    async (_sheet, _page, signal) => {
      calls++;
      expect(signal).toBe(controller.signal);
      await new Promise<void>((_resolve, reject) => {
        signal!.addEventListener(
          'abort',
          () => {
            setTimeout(() => {
              stopped = true;
              reject(signal!.reason);
            }, 0);
          },
          { once: true },
        );
        started();
      });
    },
    controller.signal,
  );
  await entered;
  controller.abort();
  expect(stopped).toBe(false);
  await expect(operation).rejects.toMatchObject({ name: 'AbortError' });
  expect(stopped).toBe(true);
  expect(calls).toBe(1);
});

it('stages a lossless downgrade and refuses rules before any output', async () => {
  const source = fixture();
  const input = upgradeValidationFile(source);
  const staged: unknown[] = [];
  const write = async (_sheet: number, page: unknown) => {
    staged.push(page);
  };
  const metadata = await stageValidationDowngrade(input, write);
  expect(metadata).toEqual({
    ...source,
    sheets: source.sheets.map(({ blocks: _blocks, ...sheet }) => sheet),
  });
  expect(staged).toEqual(source.sheets[0].blocks);
  input.sheets[1].validation.push({
    id: 'later-sheet',
    range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    rule: { kind: 'list', values: [1], allowBlank: false, policy: 'warn' },
  });
  staged.length = 0;
  await expect(stageValidationDowngrade(input, write)).rejects.toThrow('cannot be represented');
  expect(staged).toEqual([]);
});

it('stages each page once and returns rule-free metadata without copying pages', async () => {
  const source = fixture();
  const staged: unknown[] = [];
  const metadata = await stageValidationUpgrade(source, async (sheetIndex, block) => {
    expect(sheetIndex).toBe(0);
    staged.push(block);
  });
  expect(staged).toEqual(source.sheets[0].blocks);
  expect(staged[0]).toBe(source.sheets[0].blocks[0]);
  expect(metadata.schemaVersion).toBe(4);
  expect(metadata.sheets.every((sheet) => !('blocks' in sheet) && !sheet.validation.length)).toBe(
    true,
  );
  expect(metadata.sheets[0].rows).not.toBe(source.sheets[0].rows);
});

it('withholds publication metadata on cancellation or a malformed late page', async () => {
  const source = fixture();
  const controller = new AbortController();
  let count = 0;
  await expect(
    stageValidationUpgrade(
      source,
      async () => {
        count++;
        controller.abort();
      },
      controller.signal,
    ),
  ).rejects.toHaveProperty('name', 'AbortError');
  expect(count).toBe(1);
  source.sheets[0].blocks.push(source.sheets[0].blocks[0]);
  count = 0;
  await expect(
    stageValidationUpgrade(source, async () => {
      count++;
    }),
  ).rejects.toThrow('Duplicate page');
  expect(count).toBe(2);
});

function fixture() {
  const snapshot = createSnapshot({
    sheets: [
      { name: 'Data', rows: 80, columns: 4 },
      { name: 'Second', rows: 3, columns: 2 },
    ],
  });
  const sheet = snapshot.sheets[0];
  snapshot.revision = 7;
  snapshot.dateSystem = '1904';
  snapshot.styles.style1 = { bold: true, numberFormat: '0.00' };
  sheet.rows['0'] = { size: 42, hidden: true };
  sheet.freeze = { rows: 1, columns: 0 };
  sheet.cells['0:0'] = {
    rowId: 0,
    columnId: 0,
    input: { type: 'number', value: 12 },
    styleId: 'style1',
    note: 'note',
  };
  sheet.cells['64:1'] = { rowId: 64, columnId: 1, input: { type: 'formula', expression: 'A1*2' } };
  snapshot.extensions['fixture'] = { source: 'test' };
  return encodeWorkbookFile(snapshot);
}

it('round-trips a rule-free v3 file without changing content and old readers reject v4', () => {
  const source = fixture();
  const copy = structuredClone(source);
  const upgraded = upgradeValidationFile(source);
  expect(source).toEqual(copy);
  expect(upgraded.schemaVersion).toBe(4);
  expect(upgraded.sheets.map((sheet) => sheet.validation)).toEqual([[], []]);
  expect(() => validateWorkbookFile(upgraded)).toThrow('unsupported');
  const restored = downgradeValidationFile(upgraded);
  expect(restored).toEqual(source);
  restored.sheets[0].name = 'Modified';
  restored.sheets[0].rows.order[0][1] = 1;
  expect(source.sheets[0].name).toBe('Data');
  expect(upgraded.sheets[0].rows.order).toEqual(source.sheets[0].rows.order);
});

it('rejects invalid nested page objects before cloning the upgrade input', () => {
  const source = fixture();
  const page = source.sheets[0].blocks[0];
  // Define an own property to model an untrusted parsed workbook.
  Object.defineProperty(page.data, '__proto__', { value: {}, enumerable: true });
  expect(() => upgradeValidationFile(source)).toThrow('Unsafe object key');
  const invalid = fixture();
  Reflect.set(invalid.sheets[0].blocks, '0', null);
  expect(() => upgradeValidationFile(invalid)).toThrow('Invalid data block');
  const cyclic = fixture();
  const cyclicPage = cyclic.sheets[0].blocks[0];
  Reflect.set(cyclicPage.data, 'cycle', cyclicPage);
  expect(() => upgradeValidationFile(cyclic)).toThrow('Snapshot is too complex');
});

it('checks rule bounds against the total visible and hidden identities across axis runs', () => {
  const file = upgradeValidationFile(fixture());
  file.sheets[1].rows = {
    nextId: 12,
    order: [
      [8, 2, -1],
      [11, 1, 1],
    ],
    meta: { '11': { hidden: true } },
  };
  file.sheets[1].validation.push({
    id: 'last-row',
    range: { startRow: 2, endRow: 3, startColumn: 0, endColumn: 1 },
    rule: { kind: 'number', operator: 'gte', first: 1, allowBlank: false, policy: 'reject' },
  });
  expect(() => downgradeValidationFile(file)).toThrow('cannot be represented');
  file.sheets[1].validation[0].range.endRow = 4;
  expect(() => downgradeValidationFile(file)).toThrow('range');
});

it('rejects lossy downgrades and malformed validation data without mutating input', () => {
  const file = upgradeValidationFile(fixture());
  file.sheets[0].validation.push({
    id: 'required',
    range: { startRow: 1, endRow: 3, startColumn: 0, endColumn: 1 },
    rule: { kind: 'number', operator: 'gte', first: 1, allowBlank: false, policy: 'reject' },
  });
  const before = structuredClone(file);
  expect(() => downgradeValidationFile(file)).toThrow('cannot be represented');
  expect(file).toEqual(before);
  const invalid = structuredClone(file);
  invalid.sheets[0].validation[0].range.endRow = 81;
  expect(() => downgradeValidationFile(invalid)).toThrow('range');
  const missing = structuredClone(file) as any;
  delete missing.sheets[0].validation;
  expect(() => downgradeValidationFile(missing)).toThrow('Missing validation');
  const unknown = structuredClone(file) as any;
  unknown.extra = true;
  expect(() => downgradeValidationFile(unknown)).toThrow();
});

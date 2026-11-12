import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { encodeWorkbookFile, validateWorkbookFile } from '../packages/core/src/file-codec.js';
import {
  downgradeConditionalFile,
  downgradeValidationFile,
  stageConditionalDowngrade,
  stageConditionalFile,
  stageConditionalUpgrade,
  upgradeConditionalFile,
  upgradeValidationFile,
} from '../packages/core/src/validation-file-migration.js';

function fixture() {
  const snapshot = createSnapshot({
    sheets: [
      { name: 'Data', rows: 80, columns: 4 },
      { name: 'Other', rows: 10, columns: 3 },
    ],
  });
  snapshot.revision = 3;
  snapshot.sheets[0].cells['0:0'] = {
    rowId: 0,
    columnId: 0,
    input: { type: 'number', value: 4 },
    note: 'retained',
  };
  snapshot.sheets[0].cells['64:1'] = {
    rowId: 64,
    columnId: 1,
    input: { type: 'formula', expression: 'A1*2' },
  };
  const v4 = upgradeValidationFile(encodeWorkbookFile(snapshot));
  v4.sheets[0].validation.push({
    id: 'positive',
    range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    rule: { kind: 'number', operator: 'gte', first: 0, allowBlank: false, policy: 'reject' },
  });
  return v4;
}

function rule() {
  return {
    id: 'highlight',
    range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    condition: { kind: 'formula' as const, expression: 'A1>0' },
    style: { background: '#ff0000' },
    stopIfTrue: true,
  };
}

it('round-trips rule-free formatting while preserving validation and rejecting older readers', () => {
  const source = fixture();
  const before = structuredClone(source);
  const upgraded = upgradeConditionalFile(source);
  expect(upgraded.sheets.map((sheet) => sheet.conditionalFormats)).toEqual([[], []]);
  expect(() => validateWorkbookFile(upgraded)).toThrow('unsupported');
  expect(() => downgradeValidationFile(upgraded)).toThrow('schema 4 required');
  const restored = downgradeConditionalFile(upgraded);
  expect(restored).toEqual(before);
  restored.sheets[0].validation[0].range.startRow = 1;
  expect(upgraded.sheets[0].validation[0].range.startRow).toBe(0);
  expect(source).toEqual(before);
});

it('stages an upgrade without copying cell pages and validates late rule metadata first', async () => {
  const source = fixture();
  const pages: unknown[] = [];
  const metadata = await stageConditionalUpgrade(source, async (_sheet, page) => {
    pages.push(page);
  });
  expect(pages).toEqual(source.sheets[0].blocks);
  expect(pages[0]).toBe(source.sheets[0].blocks[0]);
  expect(metadata.schemaVersion).toBe(5);
  expect(metadata.sheets.map((sheet) => sheet.conditionalFormats)).toEqual([[], []]);
  expect(metadata.sheets[0].validation).toEqual(source.sheets[0].validation);
  source.sheets[1].validation.push({
    id: 'invalid',
    range: { startRow: 0, endRow: 99, startColumn: 0, endColumn: 1 },
    rule: { kind: 'number', operator: 'eq', first: 0, allowBlank: false, policy: 'reject' },
  });
  pages.length = 0;
  await expect(
    stageConditionalUpgrade(source, async (_sheet, page) => {
      pages.push(page);
    }),
  ).rejects.toThrow('range');
  expect(pages).toEqual([]);
});

it('stages detached rule-bearing metadata and rejects lossy downgrade before writing', async () => {
  const input = upgradeConditionalFile(fixture());
  input.sheets[1].conditionalFormats.push(rule());
  const pages: unknown[] = [];
  const metadata = await stageConditionalFile(input, async (_sheet, page) => {
    pages.push(page);
  });
  expect(metadata.schemaVersion).toBe(5);
  expect(metadata.sheets[1].conditionalFormats).toEqual([rule()]);
  expect(metadata.sheets[0].validation).toEqual(input.sheets[0].validation);
  expect(metadata.sheets.every((sheet) => !('blocks' in sheet))).toBe(true);
  expect(pages).toEqual(input.sheets[0].blocks);
  input.sheets[1].conditionalFormats[0].style.background = '#0000ff';
  expect(metadata.sheets[1].conditionalFormats[0].style.background).toBe('#ff0000');
  pages.length = 0;
  await expect(
    stageConditionalDowngrade(input, async (_sheet, page) => {
      pages.push(page);
    }),
  ).rejects.toMatchObject({ code: 'UNSUPPORTED_FEATURE' });
  expect(() => downgradeConditionalFile(input)).toThrow('cannot be represented');
  expect(pages).toEqual([]);
});

it('validates all rule metadata before staging and preserves v4 validation on downgrade', async () => {
  const input = upgradeConditionalFile(fixture());
  input.sheets[1].conditionalFormats.push({
    ...rule(),
    condition: { kind: 'formula', expression: 'A1>' },
  });
  let writes = 0;
  await expect(
    stageConditionalFile(input, async () => {
      writes++;
    }),
  ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  expect(writes).toBe(0);
  input.sheets[1].conditionalFormats = [];
  const metadata = await stageConditionalDowngrade(input, async () => {
    writes++;
  });
  expect(writes).toBe(2);
  expect(metadata.schemaVersion).toBe(4);
  expect(metadata.sheets[0].validation).toEqual(input.sheets[0].validation);
});

it('withholds metadata on cancellation, sink failure and malformed late pages', async () => {
  const input = upgradeConditionalFile(fixture());
  input.sheets[0].conditionalFormats.push(rule());
  const original = structuredClone(input);
  const controller = new AbortController();
  let writes = 0;
  await expect(
    stageConditionalFile(
      input,
      async () => {
        writes++;
        controller.abort();
      },
      controller.signal,
    ),
  ).rejects.toMatchObject({ name: 'AbortError' });
  expect(writes).toBe(1);
  await expect(
    stageConditionalFile(input, async () => {
      throw new Error('sink failed');
    }),
  ).rejects.toThrow('sink failed');
  expect(input).toEqual(original);
  input.sheets[0].blocks.push(input.sheets[0].blocks[0]);
  writes = 0;
  await expect(
    stageConditionalFile(input, async () => {
      writes++;
    }),
  ).rejects.toThrow('Duplicate page');
  expect(writes).toBe(2);
});

it('waits for an aborted staging sink to stop before returning', async () => {
  const input = upgradeConditionalFile(fixture());
  input.sheets[0].conditionalFormats.push(rule());
  const controller = new AbortController();
  let started!: () => void;
  const entered = new Promise<void>((resolve) => {
    started = resolve;
  });
  let stopped = false,
    calls = 0;
  const operation = stageConditionalFile(
    input,
    async (_sheet, _page, signal) => {
      calls++;
      expect(signal).toBe(controller.signal);
      await new Promise<void>((_resolve, reject) => {
        signal!.addEventListener(
          'abort',
          () =>
            setTimeout(() => {
              stopped = true;
              reject(signal!.reason);
            }, 0),
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

it('rejects missing, unknown and invalid rule metadata', () => {
  const input = upgradeConditionalFile(fixture());
  const missing = structuredClone(input);
  delete (missing.sheets[0] as any).conditionalFormats;
  expect(() => downgradeConditionalFile(missing)).toThrow('Missing conditional');
  const unknown = structuredClone(input);
  (unknown.sheets[0] as any).unexpected = true;
  expect(() => downgradeConditionalFile(unknown)).toThrow();
  const invalid = structuredClone(input);
  invalid.sheets[0].conditionalFormats.push({ ...rule(), range: { ...rule().range, endRow: 999 } });
  expect(() => downgradeConditionalFile(invalid)).toThrow('range');
});

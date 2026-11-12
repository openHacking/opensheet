import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { Workbook } from '@opensheetjs/core';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { encodeClipboardText } from '../packages/renderer/src/clipboard-text.js';

it('pins independent frozen-pane export and LibreOffice round-trip evidence', () => {
  for (const name of [
    'xlsx-freeze-export-reader-2026-11-12.json',
    'xlsx-freeze-export-reader-2026-11-12-r2.json',
  ]) {
    const report = JSON.parse(readFileSync(new URL(name, root), 'utf8'));
    for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[])
      expect(
        createHash('sha256')
          .update(readFileSync(new URL(artifact.file, root)))
          .digest('hex'),
      ).toBe(artifact.sha256);
    const expected = [
      ['Rows', 'A5', 100],
      ['Columns', 'D1', 101],
      ["Both '中' &", 'C4', 102],
      ['Unfrozen', null, 103],
    ];
    for (const file of report.independent.files)
      expect(
        file.map((sheet: { name: string; freeze: string | null; value: number }) => [
          sheet.name,
          sheet.freeze,
          sheet.value,
        ]),
      ).toEqual(expected);
  }
});
const root = new URL('../benchmarks/', import.meta.url);
const manifest = JSON.parse(
  readFileSync(new URL('./fixtures/interop/manifest.json', import.meta.url), 'utf8'),
);
const evidence = JSON.parse(
  readFileSync(new URL('interop-reader-2026-10-23-type-probe.json', root), 'utf8'),
);

it('pins independent long-formula recalculation observations', () => {
  const report = JSON.parse(
    readFileSync(new URL('formula-length-reader-2026-11-12.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  expect(report.sizesExcludingEquals).toEqual([8189, 8190, 8191, 8192, 8193, 16383]);
  expect(report.independent.files[1].map((cell: { cached: unknown }) => cell.cached)).toEqual([
    1,
    1,
    1,
    '#N/A',
    '#N/A',
    '#N/A',
  ]);
  expect(report.exportReport.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_FORMULA_LENGTH_INTEROP', address: 'A4', count: 3 }),
  );
});

it('pins independent number-format precedence and blank-link observations', () => {
  const report = JSON.parse(
    readFileSync(new URL('number-format-precedence-2026-11-12.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  const [exported, saved] = report.independent.files;
  for (const cells of [exported, saved]) {
    expect(
      cells.slice(0, 2).map((cell: { value: number; format: string }) => [cell.value, cell.format]),
    ).toEqual([
      [0.25, '0%'],
      [0.25, '0.00'],
    ]);
    expect(cells[2].format).toBe('0.00');
  }
  expect(exported[2]).toMatchObject({
    value: 'https://example.com/',
    link: 'https://example.com/',
  });
  expect(saved[2]).toMatchObject({ value: null, link: null });
});

it('pins the independent unfrozen worksheet-view fixture', () => {
  const bytes = readFileSync(new URL('xlsx-unfrozen-view-openpyxl-2026-11-12.xlsx', root));
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(
    '2f792086133131519a50f5a08a57d5c96412cd3b391f3dc57feb4154401035e5',
  );
  const workbook = XLSX.read(bytes, { type: 'buffer', bookFiles: true });
  const result = fromSheetJS(workbook);
  expect(result.report.issues.map((issue) => issue.code)).toEqual(
    expect.arrayContaining(['WORKSHEET_SELECTION_DROPPED', 'WORKSHEET_ACTIVE_TAB_DROPPED']),
  );
});

it('pins independent literal XLSX escape text observations', () => {
  const report = JSON.parse(
    readFileSync(new URL('xlsx-escape-reader-2026-11-12.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[]) {
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  }
  expect(report.libreOfficeCsv).toBe('_x0041_,_x005F_x0041_,_x0000_,_xZZZZ_\n');
  expect(report.sheetJS).toEqual({
    sheetNames: ['A'],
    cells: ['A', 'A', '\0', '_xZZZZ_'],
    note: 'note A',
    target: 'https://example.com/A',
  });
});

it('pins independent control-character XLSX reader observations', () => {
  const report = JSON.parse(
    readFileSync(new URL('xlsx-control-reader-2026-11-12.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[]) {
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  }
  expect(report.libreOfficeCsv).toContain('before_x0000_after,before_x000d_after');
  expect(report.sheetJS.cells.slice(4)).toEqual(['before\0after', 'before\rafter']);
});

it('pins independent overflow-literal reader differences', () => {
  const report = JSON.parse(
    readFileSync(new URL('formula-reader-2026-11-11-nonfinite-1900.json', root), 'utf8'),
  );
  const fixture = readFileSync(
    new URL('./fixtures/formula-corpus-nonfinite-2026-11-11.txt', import.meta.url),
  );
  expect(createHash('sha256').update(fixture).digest('hex')).toBe(report.fixtureSha256);
  for (const [name, hash] of [
    [report.sourceWorkbook, report.sourceSha256],
    [report.recalculatedWorkbook, report.recalculatedSha256],
  ])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(name, root)))
        .digest('hex'),
    ).toBe(hash);
  const cases = report.observations.filter((entry: { tag: string }) =>
    entry.tag.startsWith('nonfinite-literal'),
  );
  expect(cases).toHaveLength(7);
  expect(cases.every((entry: { agrees: boolean }) => !entry.agrees)).toBe(true);
  expect(
    cases.every((entry: { readerFormula: string }) => entry.readerFormula.includes('INF')),
  ).toBe(true);
});

it('pins the independent 1904 date-boundary corpus and reader disagreements', () => {
  const report = JSON.parse(
    readFileSync(new URL('formula-reader-2026-11-12-1904-boundaries.json', root), 'utf8'),
  );
  const fixture = readFileSync(
    new URL('./fixtures/formula-corpus-1904-boundaries-2026-11-12.txt', import.meta.url),
  );
  expect(createHash('sha256').update(fixture).digest('hex')).toBe(report.fixtureSha256);
  for (const [name, hash] of [
    [report.sourceWorkbook, report.sourceSha256],
    [report.recalculatedWorkbook, report.recalculatedSha256],
  ])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(name, root)))
        .digest('hex'),
    ).toBe(hash);
  const observations = report.observations.filter(
    (entry: { agrees: boolean | null }) => entry.agrees !== null,
  );
  expect(observations).toHaveLength(30);
  expect(observations.filter((entry: { agrees: boolean }) => entry.agrees)).toHaveLength(28);
  expect(
    observations
      .filter((entry: { agrees: boolean }) => !entry.agrees)
      .map((entry: { expression: string; actual: unknown }) => [entry.expression, entry.actual]),
  ).toEqual([
    ['DATE(1903,12,31)', -1],
    ['YEAR(2957004)', 10000],
  ]);
});

it('pins separate AVERAGE direct and reference reader observations', () => {
  const report = JSON.parse(
    readFileSync(new URL('average-coercion-reader-2026-11-11.json', root), 'utf8'),
  );
  for (const artifact of report.artifacts) {
    const bytes = readFileSync(new URL(artifact.name, root));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(artifact.sha256);
  }
  expect(report.observations).toHaveLength(8);
  expect(report.observations.map((entry: { actual: unknown }) => entry.actual)).toEqual([
    1.5,
    1,
    { error: '#VALUE!' },
    { error: '#VALUE!' },
    { error: '#VALUE!' },
    1.5,
    2,
    1,
  ]);
  const saved = report.artifacts.find((artifact: { name: string }) =>
    artifact.name.endsWith('.recalculated.xlsx'),
  );
  const workbook = XLSX.read(readFileSync(new URL(saved.name, root)), {
    type: 'buffer',
    cellFormula: true,
  });
  expect(workbook.Sheets.Data.B1.v).toBe(1.5);
  expect(workbook.Sheets.Data.B6.v).toBe(1.5);
});

it('pins independent AVERAGE invalid-text observations', () => {
  const report = JSON.parse(
    readFileSync(new URL('formula-reader-2026-11-11-average-1900.json', root), 'utf8'),
  );
  const fixture = readFileSync(
    new URL('./fixtures/formula-corpus-average-2026-11-11.txt', import.meta.url),
  );
  expect(createHash('sha256').update(fixture).digest('hex')).toBe(report.fixtureSha256);
  for (const [name, hash] of [
    [report.sourceWorkbook, report.sourceSha256],
    [report.recalculatedWorkbook, report.recalculatedSha256],
  ])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(name, root)))
        .digest('hex'),
    ).toBe(hash);
  const cases = report.observations.filter((entry: { tag: string }) =>
    entry.tag.startsWith('average-'),
  );
  expect(cases).toHaveLength(3);
  expect(cases.every((entry: { agrees: boolean }) => entry.agrees)).toBe(true);
});

it('pins independent MIN/MAX coercion observations and reader disagreements', () => {
  const report = JSON.parse(
    readFileSync(new URL('formula-reader-2026-11-11-minmax-1900.json', root), 'utf8'),
  );
  const fixture = readFileSync(
    new URL('./fixtures/formula-corpus-minmax-2026-11-11.txt', import.meta.url),
  );
  expect(createHash('sha256').update(fixture).digest('hex')).toBe(report.fixtureSha256);
  for (const [name, hash] of [
    [report.sourceWorkbook, report.sourceSha256],
    [report.recalculatedWorkbook, report.recalculatedSha256],
  ]) {
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(name, root)))
        .digest('hex'),
    ).toBe(hash);
  }
  const cases = report.observations.filter((entry: { tag: string }) =>
    [
      'direct-boolean',
      'direct-numeric-text',
      'direct-nonnumeric-text',
      'direct-empty-text',
      'referenced-boolean',
      'referenced-numeric-text',
      'computed-boolean',
    ].includes(entry.tag),
  );
  expect(cases).toHaveLength(10);
  expect(
    cases
      .filter((entry: { agrees: boolean }) => !entry.agrees)
      .map((entry: { expression: string }) => entry.expression),
  ).toEqual(['MIN("7",9)', 'MIN(A1,2)', 'MAX("7",2)']);
});

it('pins independent SUM coercion observations and reader disagreements', () => {
  const report = JSON.parse(
    readFileSync(new URL('formula-reader-2026-11-11-sum-1900.json', root), 'utf8'),
  );
  const fixture = readFileSync(
    new URL('./fixtures/formula-corpus-sum-2026-11-11.txt', import.meta.url),
  );
  expect(createHash('sha256').update(fixture).digest('hex')).toBe(report.fixtureSha256);
  for (const [name, hash] of [
    [report.sourceWorkbook, report.sourceSha256],
    [report.recalculatedWorkbook, report.recalculatedSha256],
  ])
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(name, root)))
        .digest('hex'),
    ).toBe(hash);
  const cases = report.observations.filter((entry: { tag: string }) =>
    entry.tag.startsWith('sum-'),
  );
  expect(cases).toHaveLength(7);
  expect(
    cases
      .filter((entry: { agrees: boolean }) => !entry.agrees)
      .map((entry: { expression: string }) => entry.expression),
  ).toEqual(['SUM("5",15,TRUE)', 'SUM("  +1.25e2  ",2)', 'SUM(A1,A2,2)']);
});

it('pins the independently recalculated deep-formula XLSX observation', () => {
  const report = JSON.parse(
    readFileSync(new URL('xlsx-limit-reader-2026-11-11.json', root), 'utf8'),
  );
  for (const artifact of report.artifacts) {
    const bytes = readFileSync(new URL(artifact.name, root));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(artifact.sha256);
  }
  expect(report.agrees).toBe(true);
  expect(report.report.issues).toContainEqual(
    expect.objectContaining({ code: 'RECALC_REQUIRED', count: 129 }),
  );
  expect(report.observations).toHaveLength(130);
  for (const [index, cell] of report.observations.entries()) expect(cell.value).toBe(index + 1);
  const saved = report.artifacts.find((entry: { name: string }) =>
    entry.name.endsWith('.recalculated.xlsx'),
  );
  const workbook = XLSX.read(readFileSync(new URL(saved.name, root)), {
    type: 'buffer',
    cellFormula: true,
  });
  expect(workbook.Sheets.Data.A130).toMatchObject({ f: 'A129+1', v: 130 });
});

it('pins the independent TSV file conversion observation without asserting clipboard behavior', () => {
  const report = JSON.parse(
    readFileSync(new URL('clipboard-text-reader-2026-11-11.json', root), 'utf8'),
  );
  for (const entry of report.cases) {
    for (const artifact of entry.artifacts) {
      const bytes = readFileSync(new URL(artifact.name, root));
      expect(createHash('sha256').update(bytes).digest('hex')).toBe(artifact.sha256);
    }
    const xlsx = entry.artifacts.find((artifact: { name: string }) =>
      artifact.name.endsWith('.xlsx'),
    );
    const workbook = XLSX.read(readFileSync(new URL(xlsx.name, root)), { type: 'buffer' });
    expect(workbook.Sheets[workbook.SheetNames[0]]['!ref']).toBe(entry.range);
  }
  const blank = report.cases.find((entry: { name: string }) => entry.name === 'blank');
  const control = report.cases.find((entry: { name: string }) => entry.name === 'control');
  expect(blank.inputText).toBe(encodeClipboardText([['first'], [''], ['']]));
  expect(blank).toMatchObject({ range: 'A1', csvText: 'first\n' });
  expect(control).toMatchObject({ range: 'A1:A3', csvText: 'first\n\nlast\n' });
});

it('pins blank-link output and independent reader round-trip loss', () => {
  const report = JSON.parse(
    readFileSync(new URL('annotation-interop-2026-11-11-blank-links.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[]) {
    const bytes = readFileSync(new URL(artifact.file, root));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(artifact.sha256);
  }
  for (const cell of ['A6', 'A7']) {
    const comparison = report.comparisons.find((entry: { cell: string }) => entry.cell === cell);
    expect(comparison.agrees).toBe(false);
    expect(comparison.emitted).toMatchObject({
      input: { type: 'blank' },
      link: { target: 'https://example.com/' },
    });
    expect(comparison.reopened).toMatchObject({ input: { type: 'blank' }, link: null });
  }
});

it('pins empty-text export and independent save behavior', async () => {
  const report = JSON.parse(
    readFileSync(new URL('empty-text-interop-2026-11-11.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[]) {
    const bytes = readFileSync(new URL(artifact.file, root));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(artifact.sha256);
  }
  expect(report.strictCode).toBe('UNSUPPORTED_FEATURE');
  expect(report.exportReport.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ code: 'EMPTY_TEXT_INTEROP', count: 1, address: 'A1' }),
    ]),
  );
  expect(report.before.emptyCell).toMatchObject({ t: 's', v: '' });
  expect(report.after.emptyCell).toBeNull();
  expect(report.after.countFormula).toMatchObject({ f: 'COUNTA(A1:A2)', v: 1 });
  const saved = readFileSync(new URL(report.artifacts.saved.file, root));
  const imported = fromSheetJS(XLSX.read(saved, { type: 'buffer' }));
  const book = await new Workbook(imported.file).ready();
  try {
    const sheet = book.getSheets()[0];
    expect(await sheet.range('A1:A2').getValues()).toEqual([[null], ['x']]);
    expect(await sheet.range('B1').getValues()).toEqual([[1]]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('pins independently produced workbook evidence and keeps known differences explicit', () => {
  const fixtureSource = readFileSync(new URL('./fixtures/interop/basic.fods', import.meta.url));
  expect(createHash('sha256').update(fixtureSource).digest('hex')).toBe(evidence.sourceSha256);
  for (const artifact of Object.values(evidence.artifacts) as { file: string; sha256: string }[]) {
    const bytes = readFileSync(new URL(artifact.file, root));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(artifact.sha256);
  }
  expect(
    evidence.assertions
      .filter((entry: { agrees: boolean }) => !entry.agrees)
      .map((entry: { cell: string }) => entry.cell),
  ).toEqual(['Data!C1', 'Data!G1']);
  expect(evidence.reopened.cells['Data!C1']).toMatchObject({ formula: 'TRUE()', value: 1 });
  expect(evidence.reopened.cells['Data!G1'].value).toBe(8);
  for (const snapshot of [evidence.authored, evidence.reopened]) {
    expect(snapshot.sheetNames).toEqual(['Data', 'Other']);
    expect(snapshot.cells['Data!C2']).toMatchObject({
      note: 'Public note',
      link: 'https://example.com/source',
    });
    expect(snapshot.merges).toEqual([{ s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }]);
    expect(snapshot.rows[0].hpt).toBe(32);
    expect(snapshot.rows[3].hidden).toBe(true);
  }
});

it('imports LibreOffice-authored types, notes, formats and recalculates supported formulas after edits', async () => {
  const bytes = readFileSync(new URL(evidence.artifacts.authored.file, root));
  const source = XLSX.read(bytes, { type: 'buffer', cellStyles: true, cellNF: true, xlfn: true });
  for (const [key, expected] of Object.entries(manifest.fixtures[0].expectedValues)) {
    const [sheet, cell] = key.split('!');
    expect(source.Sheets[sheet][cell].v, key).toEqual(expected);
  }
  const imported = fromSheetJS(source);
  for (const code of manifest.fixtures[0].expectedImportIssueCodes)
    expect(
      imported.report.issues.some((issue) => issue.code === code),
      code,
    ).toBe(true);
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrow();
  const book = await new Workbook(imported.file).ready();
  try {
    const sheet = book.getSheets()[0];
    expect(await sheet.range('A1:C1').getValues()).toEqual([['0001', 42, true]]);
    expect(await sheet.range('D1:E1').getDisplayValues()).toEqual([['2024-01-02', '12.5%']]);
    expect(await sheet.range('C2').getNotes()).toEqual([['Public note']]);
    expect((await sheet.range('C2').getLinks())[0][0]?.target).toBe('https://example.com/source');
    await sheet.range('B2').setValues([[9]]);
    expect(await sheet.range('F1').getValues()).toEqual([[51]]);
    expect(await book.getSheets()[1].range('A1').getValues()).toEqual([[84]]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('pins LibreOffice error-cell evidence and reimports the saved workbook', async () => {
  const report = JSON.parse(
    readFileSync(new URL('error-interop-2026-10-31-verified.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[]) {
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  }
  expect(report.assertions.map((entry: { agrees: boolean }) => entry.agrees)).toEqual([
    true,
    true,
    true,
    true,
  ]);
  expect(report.strictCode).toBe('UNSUPPORTED_FEATURE');
  expect(report.exportReport.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ code: 'ERROR_CODE_APPROXIMATED', count: 2 }),
    ]),
  );
  const saved = readFileSync(new URL(report.artifacts.saved.file, root));
  const imported = fromSheetJS(XLSX.read(saved, { type: 'buffer' }));
  const book = await new Workbook(imported.file).ready();
  try {
    expect(await book.getSheets()[0].range('A1:A4').getValues()).toEqual([
      [{ error: '#DIV/0!' }],
      [{ error: '#N/A' }],
      [{ error: '#VALUE!' }],
      [{ error: '#VALUE!' }],
    ]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('pins LibreOffice annotation round-trip evidence including empty notes', async () => {
  const report = JSON.parse(
    readFileSync(new URL('annotation-interop-2026-10-31.json', root), 'utf8'),
  );
  for (const artifact of Object.values(report.artifacts) as { file: string; sha256: string }[]) {
    expect(
      createHash('sha256')
        .update(readFileSync(new URL(artifact.file, root)))
        .digest('hex'),
    ).toBe(artifact.sha256);
  }
  expect(report.comparisons.map((entry: { agrees: boolean }) => entry.agrees)).toEqual([
    true,
    true,
    true,
    true,
    true,
  ]);
  expect(report.strictCode).toBe('UNSUPPORTED_FEATURE');
  expect(report.exportReport.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ code: 'EMPTY_LINK_TOOLTIP_DROPPED', address: 'A5' }),
    ]),
  );
  const saved = XLSX.read(readFileSync(new URL(report.artifacts.saved.file, root)), {
    type: 'buffer',
    sheetStubs: true,
  });
  const book = await new Workbook(fromSheetJS(saved).file).ready();
  try {
    const sheet = book.getSheets()[0];
    expect(await sheet.range('A1:A4').getNotes()).toEqual([[''], [''], ['hello'], ['hello']]);
    expect(await sheet.range('A1:A4').getValues()).toEqual([['value'], [null], ['value'], [null]]);
    expect(await sheet.range('A5').getLinks()).toEqual([[{ target: 'https://example.com/' }]]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

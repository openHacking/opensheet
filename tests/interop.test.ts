import { decodeWorkbookFile } from '../packages/core/src/file-codec.js';
import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { fromSheetJS, toSheetJS } from '@opensheetjs/adapter-sheetjs';
import { createWorkbook, Workbook } from '@opensheetjs/core';
import { exportRange, parseDelimited, stringifyDelimited } from '@opensheetjs/formats';
describe('SheetJS', () => {
  it('round-trips formulas without a cached value', async () => {
    const original = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', f: '1+2' } } },
    });
    const exported = toSheetJS(original.file);
    const restored = await new Workbook(fromSheetJS(exported.workbook).file).ready();
    expect(await restored.getSheets()[0].range('A1').getValues()).toEqual([[3]]);
    const bytes = XLSX.write(exported.workbook, { type: 'array', bookType: 'xlsx' });
    const binary = await new Workbook(
      fromSheetJS(XLSX.read(bytes, { type: 'array', sheetStubs: true })).file,
    ).ready();
    expect(await binary.getSheets()[0].range('A1').getValues()).toEqual([[3]]);
  });
  it('performs real XLSX write/read round trips with types, dates, merges, formats and sheet order', async () => {
    const b = await createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Other' }] }),
      s = b.getSheets()[0];
    await s.range('A1:D2').setValues([
      ['0001', 42, false, null],
      ['date', 45000, true, ''],
    ]);
    await s.range('B2').setStyle({ numberFormat: 'yyyy-mm-dd' });
    await s.range('A4:B4').merge();
    const output = toSheetJS(await b.toJSON());
    const bytes = XLSX.write(output.workbook, { type: 'array', bookType: 'xlsx' });
    const imported = fromSheetJS(
      XLSX.read(bytes, { type: 'array', cellNF: true, sheetStubs: true }),
    );
    const restored = await new Workbook(imported.file).ready();
    expect(restored.getSheets().map((s) => s.name)).toEqual(['Data', 'Other']);
    expect(await restored.getSheets()[0].range('A1:D2').getValues()).toEqual([
      ['0001', 42, false, null],
      ['date', 45000, true, ''],
    ]);
    expect((await restored.getCell(restored.getSheets()[0].id, 1, 1))?.numberFormat).toBe(
      'yyyy-mm-dd',
    );
    expect(decodeWorkbookFile(imported.file).sheets[0].merges).toHaveLength(1);
  });
  it('reads dense and sparse layouts identically', async () => {
    const bytes = XLSX.write(
      XLSX.utils.book_new(
        XLSX.utils.aoa_to_sheet([
          [1, 'text', false],
          ['', null, 2],
        ]),
        'Data',
      ),
      { type: 'array', bookType: 'xlsx' },
    );
    for (const dense of [true, false]) {
      const imported = fromSheetJS(XLSX.read(bytes, { type: 'array', dense }));
      expect(
        await (await new Workbook(imported.file).ready()).getSheets()[0].range('A1:C2').getValues(),
      ).toEqual([
        [1, 'text', false],
        ['', null, 2],
      ]);
    }
  });
  it('ignores malicious !ref without allocating a rectangle', () => {
    const data = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { '!ref': 'A1:XFD1048576', A1: { t: 'n', v: 1 } } },
    });
    expect(decodeWorkbookFile(data.file).sheets[0].rowOrder).toHaveLength(100);
    expect(() =>
      fromSheetJS({ SheetNames: ['Data'], Sheets: { Data: { A100001: { t: 'n', v: 1 } } } }),
    ).toThrow();
  });
  it('converts 1904 dates without treating serials as Unix time', async () => {
    const result = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 0, z: 'yyyy-mm-dd' } } },
      Workbook: { WBProps: { date1904: true } },
    });
    const b = await new Workbook(result.file).ready();
    expect(await b.getSheets()[0].range('A1').getDisplayValues()).toEqual([['1904-01-01']]);
  });
  it('surfaces loss and supports strict mode', () => {
    const input = {
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 's', v: 'styled', s: { font: { bold: true } } } } },
    };
    expect(fromSheetJS(input).report.summary.dropped).toBeGreaterThan(0);
    expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow();
    expect(() =>
      fromSheetJS(
        { SheetNames: ['Data'], Sheets: { Data: { A1: { t: 's', v: 'ok' } } } },
        { unsupported: 'strict' },
      ),
    ).not.toThrow();
  });
  it('enforces read-only for array formula imports', async () => {
    const { file: snapshot } = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 2, f: 'SUM(A2:A3)', F: 'A1:A1' } } },
    });
    const b = await new Workbook(snapshot).ready();
    b.setReadOnly(false);
    await expect(
      async () =>
        await b
          .getSheets()[0]
          .range('A1')
          .setValues([[3]]),
    ).rejects.toThrow(/read only/);
  });
  it('does not propagate stale formula caches', async () => {
    const { file: snapshot } = fromSheetJS({
      SheetNames: ['Data'],
      Sheets: { Data: { A1: { t: 'n', v: 123, f: 'UNSUPPORTED(1)' } } },
    });
    const b = await new Workbook(snapshot).ready(),
      s = b.getSheets()[0];
    expect((await s.range('A1').getDisplayValues())[0][0]).toBe('123 †');
    await s.range('B1').setValues([[2]]);
    expect((await s.range('A1').getDisplayValues())[0][0]).toBe('#NAME?');
    expect(toSheetJS(await b.toJSON()).workbook.Sheets.Data.A1.v).toBeUndefined();
  });
});
describe('formats', () => {
  it('parses CSV quotes, multiline fields, empty cells and CRLF', () => {
    expect(parseDelimited('"a,b",c\r\n"line\n2","a""b"\r\n')).toEqual([
      ['a,b', 'c'],
      ['line\n2', 'a"b'],
    ]);
    expect(() => parseDelimited('"unfinished')).toThrow();
  });
  it('escapes HTML and LaTeX while keeping raw values intact', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1:B2').setValues([
      ['A&B', '<script>'],
      ['50%', 'a_b'],
    ]);
    const html = exportRange(await b.toJSON(), { sheetId: s.id, format: 'html' }).text;
    expect(html).toContain('&lt;script&gt;');
    expect(html).not.toContain('<script>');
    const tex = exportRange(await b.toJSON(), { sheetId: s.id, format: 'latex' });
    expect(tex.text).toContain('A\\&B');
    expect(tex.text).toContain('50\\%');
    expect(tex.text).toContain('a\\_b');
    expect(tex.requiredPackages).toEqual(['booktabs']);
  });
  it('generates merged HTML and LaTeX', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1').setValues([['Group']]);
    await s.range('A1:B2').merge();
    const html = exportRange(await b.toJSON(), {
      sheetId: s.id,
      range: 'A1:C3',
      format: 'html',
    }).text;
    expect(html).toContain('rowspan="2" colspan="2"');
    const tex = exportRange(await b.toJSON(), {
      sheetId: s.id,
      range: 'A1:C3',
      format: 'latex',
    });
    expect(tex.text).toContain('\\multicolumn{2}{c}{\\multirow{2}{*}{Group}}');
  });
  it('protects CSV formula injection without corrupting numeric negatives', () => {
    const result = stringifyDelimited([['=HYPERLINK("bad")', -5, '+1']]);
    expect(result.sanitized).toBe(2);
    expect(result.text).toContain(',-5,');
    expect(stringifyDelimited([['=1']], ',', false).text).toBe('=1');
  });
});

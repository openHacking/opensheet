import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { FileReader } from '@opensheetjs/core';

const fixture = readFileSync('benchmarks/xlsx-link-reader-2026-11-21.saved.xlsx');
const relPath = 'xl/worksheets/_rels/sheet1.xml.rels';
const sheetPath = 'xl/worksheets/sheet1.xml';
type RawWorkbook = XLSX.WorkBook & { files?: Record<string, { content: Uint8Array }> };
const load = () => XLSX.read(fixture, { type: 'buffer', bookFiles: true }) as RawWorkbook;
const rewrite = (workbook: RawWorkbook, part: string, transform: (xml: string) => string) => {
  const item = workbook.files![part];
  const text = new TextDecoder('utf-8', { fatal: true }).decode(item.content as Uint8Array);
  const changed = transform(text);
  expect(changed).not.toBe(text);
  item.content = new TextEncoder().encode(changed);
};
const read = (workbook: RawWorkbook) => {
  const converted = fromSheetJS(workbook);
  const reader = new FileReader(converted.file);
  try {
    const sheet = converted.file.sheets[0].id;
    return { report: converted.report, target: reader.getCell(sheet, 0, 0)?.link?.target };
  } finally {
    reader.dispose();
  }
};

it('restores only targets corroborated by worksheet, relationship and SheetJS metadata', () => {
  const source = load();
  const prior = source.Sheets.Links.A1.l.Target;
  const result = read(source);
  expect(result.target).toBe('https://xn--fsqu00a.xn--0zwm56d/é/🙂?x=1&y=2#片段');
  expect(result.target).not.toBe(prior);
  expect(source.Sheets.Links.A1.l.Target).toBe(prior);
  expect(result.report.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        code: 'XLSX_LINK_URL_UTF8_RESTORED',
        count: 2,
        action: 'preserved',
      }),
    ]),
  );
});

it.each([
  [
    'missing raw parts',
    (book: RawWorkbook) => {
      delete book.files;
    },
  ],
  [
    'duplicate relationship ID',
    (book: RawWorkbook) =>
      rewrite(book, relPath, (xml) =>
        xml.replace(
          '</Relationships>',
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/伪造" TargetMode="External"/></Relationships>',
        ),
      ),
  ],
  [
    'wrong relationship mode',
    (book: RawWorkbook) =>
      rewrite(book, relPath, (xml) =>
        xml.replace('TargetMode="External"', 'TargetMode="Internal"'),
      ),
  ],
  [
    'range rather than single-cell reference',
    (book: RawWorkbook) =>
      rewrite(book, sheetPath, (xml) =>
        xml.replace('<hyperlink ref="A1"', '<hyperlink ref="A1:A2"'),
      ),
  ],
  [
    'disagreeing relationship target',
    (book: RawWorkbook) => rewrite(book, relPath, (xml) => xml.replace('/é/🙂?', '/伪造/🙂?')),
  ],
  [
    'wrong relationship namespace',
    (book: RawWorkbook) =>
      rewrite(book, relPath, (xml) =>
        xml.replace(
          'http://schemas.openxmlformats.org/package/2006/relationships',
          'https://example.com/other',
        ),
      ),
  ],
  [
    'invalid UTF-8 relationship data',
    (book: RawWorkbook) => {
      const bytes = new Uint8Array(book.files![relPath].content as Uint8Array);
      const index = bytes.indexOf(0xc3);
      expect(index).toBeGreaterThan(0);
      bytes[index] = 0xff;
      book.files![relPath].content = bytes;
    },
  ],
  [
    'relationship inspection budget exceeded',
    (book: RawWorkbook) => {
      book.files![relPath].content = new Uint8Array(4 * 1024 * 1024 + 1);
    },
  ],
  [
    'mismatched worksheet reference',
    (book: RawWorkbook) =>
      rewrite(book, sheetPath, (xml) => xml.replace('<hyperlink ref="A1"', '<hyperlink ref="A9"')),
  ],
  [
    'duplicate worksheet reference',
    (book: RawWorkbook) =>
      rewrite(book, sheetPath, (xml) =>
        xml.replace('</hyperlinks>', '<hyperlink ref="A1" r:id="rId2"/></hyperlinks>'),
      ),
  ],
  [
    'mismatched worksheet location',
    (book: RawWorkbook) =>
      rewrite(book, sheetPath, (xml) => xml.replace('location="片段"', 'location="其他"')),
  ],
  [
    'mismatched SheetJS relationship',
    (book: RawWorkbook) => {
      book.Sheets.Links.A1.l.Rel.Target = 'https://example.com/forged';
    },
  ],
  [
    'mismatched SheetJS worksheet reference',
    (book: RawWorkbook) => {
      book.Sheets.Links.A1.l.ref = 'A2';
    },
  ],
  [
    'mismatched workbook identity',
    (book: RawWorkbook) => {
      (book.Workbook!.Sheets![0] as { id?: string }).id = 'rId999';
    },
  ],
] as const)('does not guess a Unicode link when %s', (_name, mutate) => {
  const source = load();
  const original = source.Sheets.Links.A1.l.Target;
  mutate(source);
  const result = read(source);
  expect(result.target).toBe(original);
  expect(result.report.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN',
        address: 'A1',
        action: 'approximated',
      }),
    ]),
  );
});

it('accepts a link already decoded correctly when raw metadata confirms it', () => {
  const source = load();
  const link = source.Sheets.Links.A1.l;
  link.Rel.Target = 'https://xn--fsqu00a.xn--0zwm56d/é/🙂?x=1&y=2';
  link.Target = `${link.Rel.Target}#片段`;
  const result = read(source);
  expect(result.target).toBe(link.Target);
  expect(result.report.issues).not.toEqual(
    expect.arrayContaining([
      expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN', address: 'A1' }),
    ]),
  );
});

it('allows strict import when raw relationships fully restore Unicode links without other losses', () => {
  const input = XLSX.read(readFileSync('benchmarks/xlsx-link-reader-2026-11-21.baseline.xlsx'), {
    type: 'buffer',
    bookFiles: true,
    cellHTML: false,
  });
  const result = fromSheetJS(input, { unsupported: 'strict' });
  const reader = new FileReader(result.file);
  try {
    const id = result.file.sheets[0].id;
    expect(reader.getCell(id, 0, 0)?.link?.target).toBe('https://例子.测试/é/🙂?x=1&y=2#片段');
    expect(reader.getCell(id, 1, 0)?.link?.target).toBe('mailto:用户@example.com?subject=你好');
    expect(result.report.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'XLSX_LINK_URL_UTF8_RESTORED',
          action: 'preserved',
          count: 2,
        }),
      ]),
    );
  } finally {
    reader.dispose();
  }
});

it('rejects uncertain Unicode links in strict mode when raw parts are unavailable', () => {
  const source = XLSX.read(fixture, { type: 'buffer', cellHTML: false }) as RawWorkbook;
  expect(() => fromSheetJS(source, { unsupported: 'strict' })).toThrowError(
    expect.objectContaining({
      code: 'UNSUPPORTED_FEATURE',
      details: expect.objectContaining({
        issues: expect.arrayContaining([
          expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN' }),
        ]),
      }),
    }),
  );
});

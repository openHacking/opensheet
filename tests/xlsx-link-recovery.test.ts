import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { FileReader } from '@opensheetjs/core';
import { rawHyperlinkTargets } from '../packages/adapter-sheetjs/src/worksheet-losses.js';

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
    'overlapping worksheet range',
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

it('restores a corroborated Unicode hyperlink across its worksheet range', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['first'], ['second'], ['third']]);
  sheet.A1.l = { Target: 'https://example.com/路径#片段' };
  XLSX.utils.book_append_sheet(workbook, sheet, 'Data');
  const zip = XLSX.CFB.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
  });
  const part = XLSX.CFB.find(zip, '/xl/worksheets/sheet1.xml')!;
  const xml = Buffer.from(part.content!).toString('utf8');
  part.content = Buffer.from(xml.replace('ref="A1" r:id=', 'ref="A1:A3" r:id='));
  part.size = part.content.length;
  const parsed = XLSX.read(XLSX.CFB.write(zip, { type: 'buffer', fileType: 'zip' }), {
    type: 'buffer',
    bookFiles: true,
    cellHTML: false,
  });
  const original = parsed.Sheets.Data.A1.l.Target;
  const converted = fromSheetJS(parsed, { unsupported: 'strict' });
  const reader = new FileReader(converted.file);
  try {
    for (let row = 0; row < 3; row++)
      expect(reader.getCell(converted.file.sheets[0].id, row, 0)?.link?.target).toBe(
        'https://example.com/路径#片段',
      );
  } finally {
    reader.dispose();
  }
  expect(parsed.Sheets.Data.A1.l.Target).toBe(original);
  expect(converted.report.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ code: 'XLSX_LINK_URL_UTF8_RESTORED', count: 3 }),
    ]),
  );
});

it.each([16381, 16382])(
  'bounds expanded hyperlink references with %i range cells plus three other links',
  (count) => {
    const source = load();
    rewrite(source, sheetPath, (xml) =>
      xml.replace('<hyperlink ref="A1"', `<hyperlink ref="A1:A${count}"`),
    );
    const targets = rawHyperlinkTargets(source.files, source.SheetNames, source.Workbook?.Sheets);
    if (count === 16381) {
      expect(targets.get('Links')?.cells.has('A16381')).toBe(true);
      expect(targets.get('Links')?.cells.has('A2')).toBe(false);
    } else expect(targets.size).toBe(0);
  },
);

it.each(['A3:A1', 'A1:A100001', '$A$1', 'A1:ZZZZ999999999999999999'])(
  'rejects ambiguous or invalid range %s',
  (range) => {
    const source = load();
    rewrite(source, sheetPath, (xml) =>
      xml.replace('<hyperlink ref="A1"', `<hyperlink ref="${range}"`),
    );
    expect(rawHyperlinkTargets(source.files, source.SheetNames, source.Workbook?.Sheets).size).toBe(
      0,
    );
    expect(read(source).report.issues).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN' })]),
    );
  },
);

it('does not restore a range cell overlapped by a location-only hyperlink', () => {
  const source = load();
  rewrite(source, sheetPath, (xml) =>
    xml.replace('</hyperlinks>', '<hyperlink ref="A1" location="other"/></hyperlinks>'),
  );
  expect(
    rawHyperlinkTargets(source.files, source.SheetNames, source.Workbook?.Sheets)
      .get('Links')
      ?.cells.has('A1'),
  ).toBe(false);
  expect(read(source).report.issues).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN', address: 'A1' }),
    ]),
  );
});

it('restores an openpyxl range hyperlink with a locally declared relationship namespace', () => {
  const bytes = readFileSync('benchmarks/xlsx-range-link-reader-2026-11-21.xlsx');
  const workbook = XLSX.read(bytes, { type: 'buffer', bookFiles: true, cellHTML: false });
  const result = fromSheetJS(workbook);
  const reader = new FileReader(result.file);
  try {
    for (let row = 0; row < 3; row++)
      expect(reader.getCell(result.file.sheets[0].id, row, 0)?.link?.target).toBe(
        'https://example.com/路径#片段',
      );
  } finally {
    reader.dispose();
  }
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LINK_URL_UTF8_RESTORED', count: 3 }),
  );
});

it.each([
  [
    'missing relationship namespace',
    (xml: string) =>
      xml.replace(
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        '',
      ),
  ],
  [
    'incorrect local relationship namespace',
    (xml: string) =>
      xml.replace(
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        'xmlns:r="https://example.com/forged"',
      ),
  ],
  [
    'relationship namespace on a sibling',
    (xml: string) =>
      xml
        .replace(
          '<hyperlinks>',
          '<hyperlinks><unrelated xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>',
        )
        .replace(
          ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ref=',
          ' ref=',
        ),
  ],
])('does not recover when %s', (_name, mutate) => {
  const bytes = readFileSync('benchmarks/xlsx-range-link-reader-2026-11-21.xlsx');
  const workbook = XLSX.read(bytes, {
    type: 'buffer',
    bookFiles: true,
    cellHTML: false,
  }) as RawWorkbook;
  rewrite(workbook, sheetPath, mutate);
  const result = fromSheetJS(workbook);
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN', count: 3 }),
  );
  expect(result.report.issues).not.toContainEqual(
    expect.objectContaining({ code: 'XLSX_LINK_URL_UTF8_RESTORED' }),
  );
});

it('accepts the relationship namespace on the hyperlinks container', () => {
  const bytes = readFileSync('benchmarks/xlsx-range-link-reader-2026-11-21.xlsx');
  const workbook = XLSX.read(bytes, {
    type: 'buffer',
    bookFiles: true,
    cellHTML: false,
  }) as RawWorkbook;
  rewrite(workbook, sheetPath, (xml) =>
    xml
      .replace(
        '<hyperlinks>',
        '<hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      )
      .replace(
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ref=',
        ' ref=',
      ),
  );
  expect(fromSheetJS(workbook).report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LINK_URL_UTF8_RESTORED', count: 3 }),
  );
});

it('accepts a correctly decoded Unicode fragment on a verified ASCII relationship in strict mode', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['linked']]);
  sheet.A1.l = { Target: 'https://example.com/path#片段' };
  XLSX.utils.book_append_sheet(workbook, sheet, 'Data');
  const parsed = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
    bookFiles: true,
    cellHTML: false,
  }) as RawWorkbook;
  const result = fromSheetJS(parsed, { unsupported: 'strict' });
  const reader = new FileReader(result.file);
  try {
    expect(reader.getCell(result.file.sheets[0].id, 0, 0)?.link?.target).toBe(
      'https://example.com/path#片段',
    );
  } finally {
    reader.dispose();
  }
  expect(result.report.issues).not.toContainEqual(
    expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN' }),
  );
  delete parsed.files;
  expect(fromSheetJS(parsed).report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN' }),
  );
  expect(() => fromSheetJS(parsed, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it.each([
  [
    'forged fragment',
    (book: RawWorkbook) =>
      rewrite(book, sheetPath, (xml) => xml.replace('location="片段"', 'location="伪造"')),
  ],
  [
    'forged ASCII relationship',
    (book: RawWorkbook) =>
      rewrite(book, relPath, (xml) =>
        xml.replace('https://example.com/path', 'https://example.com/forged'),
      ),
  ],
  [
    'duplicate hyperlink reference',
    (book: RawWorkbook) =>
      rewrite(book, sheetPath, (xml) =>
        xml.replace(
          '</hyperlinks>',
          '<hyperlink ref="A1" r:id="rId1" location="片段"/></hyperlinks>',
        ),
      ),
  ],
  [
    'mismatched parsed relationship ID',
    (book: RawWorkbook) => {
      book.Sheets.Data.A1.l.Rel.Id = 'rId999';
    },
  ],
])('keeps the Unicode fragment warning for %s', (_name, mutate) => {
  const source = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['linked']]);
  sheet.A1.l = { Target: 'https://example.com/path#片段' };
  XLSX.utils.book_append_sheet(source, sheet, 'Data');
  const book = XLSX.read(XLSX.write(source, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
    bookFiles: true,
    cellHTML: false,
  }) as RawWorkbook;
  mutate(book);
  expect(fromSheetJS(book).report.issues).toContainEqual(
    expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN', address: 'A1' }),
  );
  expect(() => fromSheetJS(book, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it.each([false, true])(
  'handles both Unicode target and Unicode fragment in either order (%s)',
  (reverse) => {
    const source = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([['first'], ['second']]);
    const targets = ['https://example.com/路径', 'https://example.com/path#片段'];
    if (reverse) targets.reverse();
    sheet.A1.l = { Target: targets[0] };
    sheet.A2.l = { Target: targets[1] };
    XLSX.utils.book_append_sheet(source, sheet, 'Mixed');
    const parsed = XLSX.read(XLSX.write(source, { type: 'buffer', bookType: 'xlsx' }), {
      type: 'buffer',
      bookFiles: true,
      cellHTML: false,
    });
    const result = fromSheetJS(parsed, { unsupported: 'strict' });
    const reader = new FileReader(result.file);
    try {
      for (let row = 0; row < 2; row++)
        expect(reader.getCell(result.file.sheets[0].id, row, 0)?.link?.target).toBe(targets[row]);
    } finally {
      reader.dispose();
    }
    expect(result.report.issues).not.toContainEqual(
      expect.objectContaining({ code: 'XLSX_LINK_URL_IMPORT_UNCERTAIN' }),
    );
    expect(result.report.issues).toContainEqual(
      expect.objectContaining({ code: 'XLSX_LINK_URL_UTF8_RESTORED', count: 1 }),
    );
  },
);

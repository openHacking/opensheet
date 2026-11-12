import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import {
  exportRange,
  parseDelimited,
  streamExport,
  type ExportOptions,
} from '@opensheetjs/formats';
import { FormulaCalculation } from '../packages/core/src/calculation.js';

describe('streamed delimited export', () => {
  it('rejects incomplete calculated exports while preserving explicit formulas and literal text', async () => {
    let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 2 }] });
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[1]]);
      await sheet
        .range('A2:A130')
        .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 1}+1`]));
      await sheet.range('B130').setValues([['#LIMIT!']]);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      sheet = book.getSheets()[0];
      const file = await book.toJSON();
      for (const format of ['csv', 'tsv', 'html', 'markdown', 'latex'] as const) {
        const config = { sheetId: sheet.id, range: 'A130:B130', format };
        expect(() => exportRange(file, config)).toThrowError(
          expect.objectContaining({ code: 'LIMIT_EXCEEDED' }),
        );
        const report = vi.fn();
        await expect(async () => {
          for await (const _ of streamExport(book, config, { onReport: report })) {
          }
        }).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
        expect(report).not.toHaveBeenCalled();
      }
      for (const format of ['csv', 'tsv'] as const) {
        const config = {
          sheetId: sheet.id,
          range: 'A130:B130',
          format,
          options: { safe: false, formulas: 'preserve' as const },
        };
        const expected = exportRange(file, config).text;
        expect(parseDelimited(expected, format === 'csv' ? ',' : '\t')).toEqual([
          ['=A129+1', '#LIMIT!'],
        ]);
        let actual = '';
        for await (const chunk of streamExport(book, config)) actual += chunk;
        expect(actual).toBe(expected);
      }
      expect(exportRange(file, { sheetId: sheet.id, range: 'B130', format: 'csv' }).text).toBe(
        '#LIMIT!',
      );
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });

  it('fails without a report when a later stream batch contains an incomplete result', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([['first']]);
      const original = book.readRange.bind(book);
      const read = vi.spyOn(book, 'readRange').mockImplementation(async (...args) => {
        const result = await original(...args);
        if (args[1].startRow > 0) {
          const metadata = book.sheetData(sheet.id);
          result.calculated[`${metadata.rowOrder[args[1].startRow]}:${metadata.columnOrder[0]}`] = {
            error: '#LIMIT!',
          };
        }
        return result;
      });
      const report = vi.fn(),
        chunks: string[] = [];
      await expect(async () => {
        for await (const chunk of streamExport(
          book,
          { sheetId: sheet.id, range: 'A1:DX130', format: 'csv' },
          { onReport: report },
        ))
          chunks.push(chunk);
      }).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
      expect(chunks).toHaveLength(1);
      expect(chunks[0]).toMatch(/^first,/);
      expect(report).not.toHaveBeenCalled();
      read.mockRestore();
    } finally {
      await book.deleteStorage();
      await book.close();
    }
  });
  it.each(['csv', 'tsv'] as const)(
    'exports formulas only in explicitly requested raw %s mode',
    async (format) => {
      const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
      try {
        const sheet = book.getSheets()[0];
        await sheet.range('B1:C1').setFormulas([['SUM(1,2)', 'IF(TRUE,"a,b","x")']]);
        await sheet.range('B130').setFormulas([['1/0']]);
        await sheet.range('C130').setValues([['=literal']]);
        await sheet.range('D1:F1').setValues([[42, false, 'plain']]);
        await book.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [{ row: 0, column: 6, input: { type: 'error', code: '#N/A' } }],
          },
        });
        const base = { sheetId: sheet.id, range: 'B1:DX130', format };
        const file = await book.toJSON();
        const separator = format === 'csv' ? ',' : '\t';
        const normal = parseDelimited(exportRange(file, base).text, separator);
        expect(normal[0].slice(0, 2)).toEqual(['3', 'a,b']);
        expect(normal[129].slice(0, 2)).toEqual(['#DIV/0!', "'=literal"]);
        const rawValues = parseDelimited(
          exportRange(file, { ...base, options: { safe: false } }).text,
          separator,
        );
        expect(rawValues[0].slice(0, 2)).toEqual(['3', 'a,b']);
        expect(rawValues[129][1]).toBe('=literal');
        const config = { ...base, options: { safe: false, formulas: 'preserve' as const } };
        const calculate = vi.spyOn(FormulaCalculation.prototype, 'value');
        const complete = exportRange(file, config);
        expect(calculate).not.toHaveBeenCalled();
        calculate.mockRestore();
        const values = parseDelimited(complete.text, separator);
        expect(values[0].slice(0, 2)).toEqual(['=SUM(1,2)', '=IF(TRUE,"a,b","x")']);
        expect(values[0].slice(2, 6)).toEqual(['42', 'false', 'plain', '#N/A']);
        expect(values[129].slice(0, 2)).toEqual(['=1/0', '=literal']);
        expect(complete.report).toContain(
          '3 formulas exported as formulas; raw text is not escaped and spreadsheet applications may execute formulas. References are not rebased and may depend on omitted cells or sheets.',
        );
        let streamed = '';
        const report = vi.fn();
        const read = vi.spyOn(book, 'readRange');
        for await (const chunk of streamExport(book, config, { onReport: report }))
          streamed += chunk;
        expect(read).toHaveBeenCalled();
        expect(read.mock.calls.every((call) => call[2] === false)).toBe(true);
        read.mockRestore();
        expect(streamed).toBe(complete.text);
        expect(report).toHaveBeenCalledWith(
          expect.objectContaining({ sanitized: 0, revision: book.revision, preservedFormulas: 3 }),
        );
        for (const options of [
          { formulas: 'preserve' as const },
          { formulas: 'preserve' as const, safe: true },
        ]) {
          expect(() => exportRange(file, { ...base, options })).toThrow(
            'requires explicit safe: false',
          );
          await expect(streamExport(book, { ...base, options }).next()).rejects.toMatchObject({
            code: 'INVALID_ARGUMENT',
          });
        }
        expect(() =>
          exportRange(file, { ...base, options: { formulas: 'invalid' as never } }),
        ).toThrow('Unknown delimited formula mode');
      } finally {
        await book.close();
      }
    },
  );
  it('keeps the initial format, safety policy, range and report callback across batches', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([['=first']]);
      await sheet.range('A130').setValues([['=last']]);
      const range = { startRow: 0, endRow: 130, startColumn: 0, endColumn: 128 };
      const config: ExportOptions = {
        sheetId: sheet.id,
        range,
        format: 'csv',
        options: { safe: true, bom: true },
      };
      const original = vi.fn(),
        replacement = vi.fn();
      const options = { onReport: original };
      const stream = streamExport(book, config, options);
      const first = await stream.next();
      config.format = 'tsv';
      config.options!.safe = false;
      config.options!.bom = false;
      range.startRow = 50;
      options.onReport = replacement;
      const chunks = [first.value!];
      for await (const chunk of stream) chunks.push(chunk);
      const result = parseDelimited(chunks.join(''));
      expect(result).toHaveLength(130);
      expect(result[0][0]).toBe("'=first");
      expect(result[129][0]).toBe("'=last");
      expect(chunks.join('').match(/\uFEFF/g)).toHaveLength(1);
      expect(original).toHaveBeenCalledExactlyOnceWith({
        format: 'csv',
        revision: book.revision,
        sanitized: 2,
        sample: ['A1', 'A130'],
        truncated: false,
      });
      expect(replacement).not.toHaveBeenCalled();
    } finally {
      await book.close();
    }
  });
  it('reports bounded affected addresses only after a complete revision-bound stream', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('B1:B33').setValues(Array.from({ length: 33 }, () => ['=text']));
      await sheet.range('DX130').setValues([['＝text']]);
      const report = vi.fn();
      const config = { sheetId: sheet.id, range: 'B1:DX130', format: 'csv' as const };
      const stream = streamExport(book, config, { onReport: report });
      expect((await stream.next()).done).toBe(false);
      expect(report).not.toHaveBeenCalled();
      for await (const _chunk of stream) {
        /* consume the remaining stream */
      }
      expect(report).toHaveBeenCalledExactlyOnceWith({
        format: 'csv',
        revision: book.revision,
        sanitized: 34,
        sample: Array.from({ length: 32 }, (_, i) => `B${i + 1}`),
        truncated: true,
      });
      const raw = vi.fn();
      for await (const _chunk of streamExport(
        book,
        { ...config, options: { safe: false } },
        { onReport: raw },
      )) {
        /* consume */
      }
      expect(raw).toHaveBeenCalledExactlyOnceWith({
        format: 'csv',
        revision: book.revision,
        sanitized: 0,
        sample: [],
        truncated: false,
      });
      const aborted = vi.fn();
      const partial = streamExport(book, config, { onReport: aborted });
      await partial.next();
      await partial.return(undefined);
      expect(aborted).not.toHaveBeenCalled();
      const controller = new AbortController();
      const cancelled = streamExport(book, config, {
        onReport: aborted,
        signal: controller.signal,
      });
      await cancelled.next();
      controller.abort();
      await expect(cancelled.next()).rejects.toMatchObject({ name: 'AbortError' });
      expect(aborted).not.toHaveBeenCalled();
      const stale = streamExport(book, config, { onReport: aborted });
      await stale.next();
      await sheet.range('A1').setValues([[1]]);
      await expect(stale.next()).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
      expect(aborted).not.toHaveBeenCalled();
    } finally {
      await book.close();
    }
  });
  it('keeps the original cancellation signal when options are mutated', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
    try {
      const sheet = book.getSheets()[0],
        initial = new AbortController(),
        replacement = new AbortController();
      const options = { signal: initial.signal };
      const stream = streamExport(
        book,
        { sheetId: sheet.id, range: 'A1:DX130', format: 'csv' },
        options,
      );
      await stream.next();
      options.signal = replacement.signal;
      initial.abort();
      await expect(stream.next()).rejects.toMatchObject({ name: 'AbortError' });
    } finally {
      await book.close();
    }
  });
  it.each(['csv', 'tsv'] as const)(
    'preserves safe and raw prefix policy in %s exports',
    async (format) => {
      const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 1, columns: 6 }] });
      try {
        const sheet = book.getSheets()[0];
        const values = ['＝1', ' ＋1', '\ntext', '\0text', -5, 'ordinary,\t"=1'];
        await sheet.range('A1:F1').setValues([values]);
        for (const safe of [true, false]) {
          const config = { sheetId: sheet.id, range: 'A1:F1', format, options: { safe } };
          const complete = exportRange(await book.toJSON(), config);
          let streamed = '';
          for await (const chunk of streamExport(book, config)) streamed += chunk;
          expect(streamed).toBe(complete.text);
          expect(parseDelimited(streamed, format === 'csv' ? ',' : '\t')).toEqual([
            [
              ...(safe ? values.slice(0, 4).map((value) => `'${value}`) : values.slice(0, 4)),
              '-5',
              values[5],
            ],
          ]);
          expect(complete.report.some((message) => message.includes('4 text cells prefixed'))).toBe(
            safe,
          );
        }
      } finally {
        await book.close();
      }
    },
  );
  it.each(['csv', 'tsv'] as const)(
    'matches the complete %s output across range batches',
    async (format) => {
      const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
      try {
        const sheet = book.getSheets()[0];
        await sheet.range('A1:B1').setValues([['=unsafe', -5]]);
        await sheet.range('A128:B129').setValues([
          ['line\nbreak', 'a,b'],
          ['+unsafe', 'quote"value'],
        ]);
        const config = { sheetId: sheet.id, range: 'A1:DX130', format, options: { bom: true } };
        const expected = exportRange(await book.toJSON(), config).text;
        let actual = '';
        for await (const chunk of streamExport(book, config)) actual += chunk;
        expect(actual).toBe(expected);
        expect(actual.match(/\uFEFF/g)).toHaveLength(1);
      } finally {
        await book.close();
      }
    },
  );
});

describe('streamed rich export', () => {
  it.each(['markdown', 'html', 'latex'] as const)(
    'preserves %s output with a merge crossing a batch boundary',
    async (format) => {
      const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
      try {
        const sheet = book.getSheets()[0];
        await sheet.range('A1:B1').setValues([['Header', 'Second']]);
        await sheet.range('A128').setValues([['merged']]);
        await sheet.range('A128:A129').merge();
        await sheet.range('B129').setValues([['<img src=x onerror=alert(1)> & after | quote']]);
        await sheet.range('B129').setStyle({ bold: true, color: '#ff0000' });
        const config = { sheetId: sheet.id, range: 'A1:DX130', format };
        const expected = exportRange(await book.toJSON(), config).text;
        let actual = '';
        for await (const chunk of streamExport(book, config)) actual += chunk;
        expect(actual).toBe(expected);
        if (format === 'markdown') {
          expect(actual).not.toContain('<img');
          expect(actual).toContain('&lt;img src=x onerror=alert\\(1\\)&gt; &amp; after \\| quote');
        }
      } finally {
        await book.close();
      }
    },
  );
});

it('exports Markdown cell contents as literal text rather than inline markup', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Text', rows: 2, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet
      .range('A1:A2')
      .setValues([
        ['Header'],
        ['![alt](https://example.com/x.png) *bold* `code` <img src=x> a|b \\ slash\rline'],
      ]);
    const config = { sheetId: sheet.id, format: 'markdown' as const };
    const expected =
      '| Header |\n| --- |\n| \\!\\[alt\\]\\(https://example.com/x.png\\) \\*bold\\* \\`code\\` &lt;img src=x&gt; a\\|b \\\\ slash<br>line |';
    expect(exportRange(await book.toJSON(), config).text).toBe(expected);
    let actual = '';
    for await (const chunk of streamExport(book, config)) actual += chunk;
    expect(actual).toBe(expected);
  } finally {
    await book.close();
  }
});

it('exports all common HTML line endings without interpreting cell markup', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Text', rows: 2, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet
      .range('A1:A2')
      .setValues([['Heading'], ['first\rsecond\r\n<img src=x onerror="alert(1)">\nlast & end']]);
    const config = { sheetId: sheet.id, format: 'html' as const };
    const expected = exportRange(await book.toJSON(), config).text;
    expect(expected).toContain(
      'first<br>second<br>&lt;img src=x onerror=&quot;alert(1)&quot;&gt;<br>last &amp; end',
    );
    expect(expected).not.toContain('<img');
    expect(expected).not.toContain('\r');
    let streamed = '';
    for await (const chunk of streamExport(book, config)) streamed += chunk;
    expect(streamed).toBe(expected);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('rejects an HTML export when a structural edit happens after the opening chunk', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 3, columns: 2 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A2').setValues([['first'], ['second']]);
    const stream = streamExport(book, { sheetId: sheet.id, range: 'A1:B2', format: 'html' });
    expect((await stream.next()).value).toBe('<table>\n');
    await sheet.insertRows(0);
    await expect(stream.next()).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
  } finally {
    await book.close();
  }
});

it('honors cancellation before and after the HTML opening chunk', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] });
  try {
    const sheet = book.getSheets()[0],
      controller = new AbortController();
    const config = { sheetId: sheet.id, format: 'html' as const };
    const stream = streamExport(book, config, { signal: controller.signal });
    expect((await stream.next()).value).toBe('<table>\n');
    controller.abort();
    await expect(stream.next()).rejects.toMatchObject({ name: 'AbortError' });
    await expect(
      streamExport(book, config, { signal: controller.signal }).next(),
    ).rejects.toMatchObject({ name: 'AbortError' });
  } finally {
    await book.close();
  }
});

it('rejects a revision change between delimited batches', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
  try {
    const sheet = book.getSheets()[0];
    const stream = streamExport(book, { sheetId: sheet.id, range: 'A1:DX130', format: 'csv' });
    expect((await stream.next()).done).toBe(false);
    await sheet.range('A129').setValues([[42]]);
    await expect(stream.next()).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
  } finally {
    await book.close();
  }
});

it('rejects an old read response after the workbook revision changes', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] });
  try {
    const sheet = book.getSheets()[0];
    let release!: () => void, captured!: () => void;
    const waiting = new Promise<void>((resolve) => {
      release = resolve;
    });
    const ready = new Promise<void>((resolve) => {
      captured = resolve;
    });
    const read = book.readRange.bind(book);
    vi.spyOn(book, 'readRange').mockImplementation(async (...args) => {
      const result = await read(...args);
      captured();
      await waiting;
      return result;
    });
    const stream = book.streamRange(sheet.id, {
      startRow: 0,
      endRow: 2,
      startColumn: 0,
      endColumn: 2,
    });
    const result = stream.next();
    await ready;
    await sheet.range('A1').setValues([[42]]);
    release();
    await expect(result).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
  } finally {
    await book.close();
  }
});

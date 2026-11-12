import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

// Independent left-to-right oracle: no regular expressions or built-in replace operation.
function replaceLiteral(source: string, query: string, replacement: string) {
  let output = '';
  for (let at = 0; at < source.length; ) {
    if (source.startsWith(query, at)) {
      output += replacement;
      at += query.length;
    } else output += source[at++];
  }
  return output;
}

describe('literal replacement model', () => {
  it('matches an independent non-overlapping UTF-16 model over replayable generated inputs', () => {
    const text = (minLength: number, maxLength: number) =>
      fc
        .array(fc.constantFrom('a', 'b', '$', '&', "'", '😀', '\n', '\ud800'), {
          minLength,
          maxLength,
        })
        .map((parts) => parts.join(''));
    fc.assert(
      fc.property(text(0, 40), text(1, 4), text(0, 8), (source, query, replacement) => {
        const snapshot = createSnapshot({ sheets: [{ name: 'Model', rows: 1, columns: 1 }] });
        const sheet = snapshot.sheets[0];
        sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'string', value: source } };
        reduceCommand(snapshot, {
          type: 'core.cells.replace',
          payload: { sheetId: sheet.id, query, replacement },
        });
        expect(sheet.cells['0:0'].input).toEqual({
          type: 'string',
          value: replaceLiteral(source, query, replacement),
        });
      }),
      { seed: 271828, numRuns: 1000 },
    );
  });

  it('accepts exactly the text limit and rejects one unit over without changing history', async () => {
    let book = await createWorkbook();
    try {
      let sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([['a'.repeat(16384)]]);
      await book.replaceText(sheet.id, 'a', 'aa');
      expect((await sheet.range('A1').getValues())[0][0]).toBe('a'.repeat(32768));
      await book.undo();
      await sheet.range('A1').setValues([['a'.repeat(16385)]]);
      const revision = book.revision;
      await expect(book.replaceText(sheet.id, 'a', 'aa')).rejects.toMatchObject({
        code: 'LIMIT_EXCEEDED',
      });
      expect(book.revision).toBe(revision);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      sheet = book.getSheets()[0];
      expect((await sheet.range('A1').getValues())[0][0]).toBe('a'.repeat(16385));
      await book.undo();
      expect((await sheet.range('A1').getValues())[0][0]).toBe('a'.repeat(16384));
    } finally {
      await book.close();
    }
  });
});

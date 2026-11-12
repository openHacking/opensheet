import { expect, it } from 'vitest';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';

const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
const frozen = (x: number, y: number, topLeft: string) =>
  `<worksheet xmlns="${main}"><sheetViews><sheetView workbookViewId="0"><pane xSplit="${x}" ySplit="${y}" topLeftCell="${topLeft}" state="frozen"/></sheetView></sheetViews><sheetData/></worksheet>`;
const bytes = (text: string) => ({ content: new TextEncoder().encode(text) });
function workbook(first = frozen(2, 3, 'C4'), second = frozen(1, 0, 'B1')) {
  return {
    SheetNames: ['First', 'Second'],
    Sheets: {
      First: { A1: { t: 'n', v: 1 }, '!ref': 'A1' },
      Second: { A1: { t: 'n', v: 2 }, '!ref': 'A1' },
    },
    Workbook: {
      Sheets: [
        { name: 'First', id: 'r1' },
        { name: 'Second', id: 'r2' },
      ],
    },
    files: {
      'xl/workbook.xml': bytes(
        `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="First" r:id="r1"/><sheet name="Second" r:id="r2"/></sheets></workbook>`,
      ),
      'xl/_rels/workbook.xml.rels': bytes(
        `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="r2" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      ),
      'xl/worksheets/sheet1.xml': bytes(first),
      'xl/worksheets/sheet2.xml': bytes(second),
    },
  };
}

it('maps panes by workbook relationships, not ZIP part order, and expands sheet dimensions', () => {
  const input = workbook(frozen(40, 150, 'AO151'));
  const result = fromSheetJS(input, { unsupported: 'strict' });
  expect(result.file.sheets.map((sheet) => sheet.freeze)).toEqual([
    { rows: 0, columns: 1 },
    { rows: 150, columns: 40 },
  ]);
  expect(result.file.sheets[1].rows.order.reduce((sum, [, count]) => sum + count, 0)).toBe(151);
  expect(result.file.sheets[1].columns.order.reduce((sum, [, count]) => sum + count, 0)).toBe(41);
  expect(result.report.issues.some((issue) => issue.code === 'WORKSHEET_PANE_UNSUPPORTED')).toBe(
    false,
  );
});

it('maps frozen panes from BOM-less UTF-16 workbook and worksheet parts', () => {
  for (const encoding of ['le', 'be']) {
    const input = workbook();
    for (const part of Object.values(input.files)) {
      const encoded = Buffer.from(new TextDecoder().decode(part.content), 'utf16le');
      if (encoding === 'be') encoded.swap16();
      part.content = encoded;
    }
    expect(
      fromSheetJS(input, { unsupported: 'strict' }).file.sheets.map((sheet) => sheet.freeze),
    ).toEqual([
      { rows: 0, columns: 1 },
      { rows: 3, columns: 2 },
    ]);
  }
});

it('reports ambiguous sheet mapping rather than applying a pane to the wrong sheet', () => {
  const input = workbook();
  input.Workbook.Sheets[0].id = 'r2';
  const result = fromSheetJS(input);
  expect(result.file.sheets.map((sheet) => sheet.freeze)).toEqual([
    { rows: 0, columns: 0 },
    { rows: 0, columns: 0 },
  ]);
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_PANE_UNSUPPORTED', count: 2 }),
  );
});

it('does not import panes with unsupported viewport offsets, namespaces or multiple views', () => {
  for (const xml of [
    frozen(2, 3, 'C10'),
    frozen(2, 3, 'C4').replace('workbookViewId="0"', 'workbookViewId="1"'),
    frozen(2, 3, 'C4').replace(
      '<sheetView workbookViewId="0">',
      '<sheetView xmlns="urn:other" workbookViewId="0">',
    ),
    frozen(2, 3, 'C4').replace('</sheetViews>', '<sheetView workbookViewId="0"/></sheetViews>'),
    `<worksheet xmlns="${main}"><extLst><pane xSplit="2" ySplit="3" topLeftCell="C4" state="frozen"/></extLst></worksheet>`,
  ]) {
    const result = fromSheetJS(
      workbook(xml, '<worksheet xmlns="' + main + '"><sheetData/></worksheet>'),
    );
    expect(result.file.sheets[1].freeze).toEqual({ rows: 0, columns: 0 });
    expect(result.report.issues).toContainEqual(
      expect.objectContaining({ code: 'WORKSHEET_PANE_UNSUPPORTED', count: 1 }),
    );
  }
});

it('imports neutral LibreOffice view attributes while reporting selected tabs', () => {
  const base = frozen(2, 3, 'C4');
  const view =
    'workbookViewId="0" showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100"';
  const input = workbook(
    base.replace('workbookViewId="0"', view),
    '<worksheet xmlns="' + main + '"><sheetData/></worksheet>',
  );
  const result = fromSheetJS(input);
  expect(result.file.sheets[1].freeze).toEqual({ rows: 3, columns: 2 });
  expect(result.report.issues).not.toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_PANE_UNSUPPORTED' }),
  );
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_ACTIVE_TAB_DROPPED', count: 1 }),
  );
  expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
  for (const [before, after] of [
    ['showGridLines="true"', 'showGridLines="false"'],
    ['rightToLeft="false"', 'rightToLeft="true"'],
    ['zoomScale="100"', 'zoomScale="125"'],
    ['topLeftCell="A1"', 'topLeftCell="B2"'],
    ['workbookViewId="0"', 'workbookViewId="0" constructor="x"'],
  ]) {
    const invalid = workbook(
      base.replace('workbookViewId="0"', view.replace(before, after)),
      '<worksheet xmlns="' + main + '"><sheetData/></worksheet>',
    );
    const rejected = fromSheetJS(invalid);
    expect(rejected.file.sheets[1].freeze).toEqual({ rows: 0, columns: 0 });
    expect(rejected.report.issues).toContainEqual(
      expect.objectContaining({ code: 'WORKSHEET_PANE_UNSUPPORTED' }),
    );
  }
});

it('reports active tabs and selections even when no pane is present', () => {
  const xml = `<worksheet xmlns="${main}"><sheetViews><sheetView workbookViewId="0" tabSelected="true"><selection activeCell="D8" sqref="D8"/></sheetView></sheetViews><sheetData/></worksheet>`;
  const input = workbook(xml, '<worksheet xmlns="' + main + '"><sheetData/></worksheet>');
  const result = fromSheetJS(input);
  expect(result.file.sheets[1].freeze).toEqual({ rows: 0, columns: 0 });
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_ACTIVE_TAB_DROPPED', count: 1 }),
  );
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_SELECTION_DROPPED', count: 1 }),
  );
  expect(result.report.issues).not.toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_PANE_UNSUPPORTED' }),
  );
  expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
});

it('refuses duplicate relationship identities, including non-worksheet relationships', () => {
  const input = workbook();
  input.files['xl/_rels/workbook.xml.rels'] = bytes(
    `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/styles" Target="styles.xml"/><Relationship Id="r1" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="r2" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  const result = fromSheetJS(input);
  expect(result.file.sheets[0].freeze).toEqual({ rows: 0, columns: 0 });
  expect(result.file.sheets[1].freeze).toEqual({ rows: 3, columns: 2 });
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_PANE_UNSUPPORTED', count: 1 }),
  );
});

it('preserves supported freeze counts but reports a non-default selection', () => {
  const xml = frozen(2, 3, 'C4').replace(
    '</sheetView>',
    '<selection activeCell="D8" sqref="D8"/></sheetView>',
  );
  const result = fromSheetJS(
    workbook(xml, '<worksheet xmlns="' + main + '"><sheetData/></worksheet>'),
  );
  expect(result.file.sheets[1].freeze).toEqual({ rows: 3, columns: 2 });
  expect(result.report.issues).toContainEqual(
    expect.objectContaining({ code: 'WORKSHEET_SELECTION_DROPPED', count: 1 }),
  );
  expect(() =>
    fromSheetJS(workbook(xml, '<worksheet xmlns="' + main + '"><sheetData/></worksheet>'), {
      unsupported: 'strict',
    }),
  ).toThrow('Strict compatibility check failed');
});

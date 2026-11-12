import { build } from 'esbuild';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import fc from 'fast-check';
import { expect, it } from 'vitest';

it('matches the former XML tag grammar on generated worksheet structures', async () => {
  const file = resolve('packages/adapter-sheetjs/src/worksheet-losses.ts');
  const source = readFileSync(file, 'utf8');
  const start = source.indexOf('      // Avoid a regular-expression match allocation');
  const marker = '        const nodeName = tag.slice(nameStart, nameEnd);';
  const end = source.indexOf(marker, start);
  expect(start).toBeGreaterThanOrEqual(0);
  expect(end).toBeGreaterThan(start);
  const previous =
    source.slice(0, start) +
    String.raw`      const node = inspectTree ? /^<(\/?)([\w.:-]+)(?=[\s/>])/.exec(tag) : null;
      if (node) {
        const closing = node[1] === '/';
        const nodeName = node[2];` +
    source.slice(end + marker.length);
  const modules: Array<{ reportWorksheetLosses: (...args: any[]) => Map<string, unknown> }> = [];
  for (const contents of [previous, source]) {
    const bundle = await build({
      entryPoints: [file],
      bundle: true,
      platform: 'node',
      format: 'esm',
      write: false,
      plugins: [
        {
          name: 'variant',
          setup(builder) {
            builder.onLoad({ filter: /worksheet-losses\.ts$/ }, () => ({
              contents,
              loader: 'ts',
              resolveDir: resolve('packages/adapter-sheetjs/src'),
            }));
          },
        },
      ],
    });
    modules.push(
      await import(
        `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].text).toString('base64')}`
      ),
    );
  }
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const encode = (text: string) => ({ content: new TextEncoder().encode(text) });
  const noise = fc.constantFrom(
    '<x/>',
    '<x></x>',
    '<x:y a="1"/>',
    '<x.name/>',
    '<x-name/>',
    '<_x/>',
    '<é/>',
    '<bad!/>',
    '</orphan>',
    '<!-- <pane/> -->',
    '<![CDATA[<selection/>]]>',
  );
  const views = fc.constantFrom(
    '<sheetView workbookViewId="0" tabSelected="true"><selection activeCell="D8" sqref="D8"/></sheetView>',
    '<sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" state="frozen"/></sheetView>',
    '<sheetView workbookViewId="0" xmlns="urn:other"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView>',
    '<sheetView workbookViewId="1" tabSelected="1"/>',
  );
  fc.assert(
    fc.property(
      fc.array(noise, { maxLength: 12 }),
      views,
      fc.array(noise, { maxLength: 12 }),
      (before, view, after) => {
        const xml = `<worksheet xmlns="${main}">${before.join('')}<sheetViews>${view}</sheetViews>${after.join('')}<sheetData/></worksheet>`;
        const files = {
          'xl/workbook.xml': encode(
            `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="S" r:id="r1"/></sheets></workbook>`,
          ),
          'xl/_rels/workbook.xml.rels': encode(
            `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
          ),
          'xl/worksheets/sheet1.xml': encode(xml),
        };
        const run = (module: any) => {
          const report = {
            adapterVersion: 'test',
            sourceFormat: 'SheetJS',
            summary: { exact: 0, approximated: 0, dropped: 0, blocked: 0 },
            issues: [],
            omittedIssueOccurrences: 0,
          };
          try {
            return {
              freeze: [
                ...module.reportWorksheetLosses(files, report, ['S'], [{ name: 'S', id: 'r1' }]),
              ],
              report,
            };
          } catch (error) {
            return { error: (error as Error).message, code: (error as { code?: string }).code };
          }
        };
        expect(run(modules[1])).toEqual(run(modules[0]));
      },
    ),
    { seed: 20261112, numRuns: 1000 },
  );
});

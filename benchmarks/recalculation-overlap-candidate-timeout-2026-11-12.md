# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: recalculation-overlap.spec.ts >> overlapping formulas preserve results and history with cache 0
- Location: tests/e2e/recalculation-overlap.spec.ts:4:3

# Error details

```
Test timeout of 30000ms exceeded.
```

```
Error: page.evaluate: Test timeout of 30000ms exceeded.
```

# Page snapshot

```yaml
- generic [ref=e2]:
  - complementary [ref=e3]:
    - link "OpenSheet." [ref=e4]:
      - /url: "#budget"
    - generic [ref=e6]:
      - generic [ref=e7]: O
      - generic [ref=e8]:
        - text: Open workspace
        - generic [ref=e9]: Open-source workspace
    - generic [ref=e13]: WORKSPACE
    - button "Spreadsheet 1" [ref=e14] [cursor=pointer]:
      - text: Spreadsheet
      - generic [ref=e17]: "1"
    - generic [ref=e18]: DEVELOPER TOOLS
    - button "API quick start" [ref=e19] [cursor=pointer]
    - button "Plugin example" [ref=e23] [cursor=pointer]
    - generic [ref=e27]:
      - generic [ref=e28]: Your data stays yours.
      - paragraph [ref=e32]: Files are processed on your device.No account. No uploads.
      - generic [ref=e33]:
        - text: Open source
        - generic [ref=e35]: v0.2.0-dev.0
  - generic [ref=e36]:
    - banner [ref=e37]:
      - generic [ref=e38]: Workspace / Playground
      - generic [ref=e39]:
        - generic [ref=e40]: Runs locally
        - button "Keyboard shortcuts" [ref=e42] [cursor=pointer]
        - generic [ref=e46]: OS
    - main [ref=e47]:
      - generic [ref=e48]:
        - region "Demo controls" [ref=e49]:
          - generic [ref=e50]:
            - generic [ref=e51]: DEMO
            - combobox "Choose demo" [ref=e52]:
              - option "Performance Lab"
              - option "Budget Simulator" [selected]
              - option "Sales Dashboard"
              - option "Project Planner"
              - option "Table to Code"
          - generic [ref=e53]:
            - generic [ref=e54]: 01 / PLAN
            - heading "Budget Simulator" [level=1] [ref=e55]
            - paragraph [ref=e56]: Turn a what-if into a clear plan. Edit prices and watch the numbers move.
          - button "Reset demo" [ref=e57] [cursor=pointer]
        - generic [ref=e61]:
          - region "Spreadsheet editor" [ref=e62]:
            - generic [ref=e63]:
              - generic [ref=e69]:
                - textbox "Workbook name" [ref=e70]: Launch budget
                - status [ref=e73]: Saved locally
              - generic [ref=e74]:
                - button "Import file" [ref=e75] [cursor=pointer]
                - button "Save" [ref=e80] [cursor=pointer]
                - button "Export JSON" [ref=e81] [cursor=pointer]
                - button "Export XLSX" [ref=e82] [cursor=pointer]
            - generic [ref=e87]:
              - generic [ref=e88]:
                - textbox "Find in sheet" [ref=e92]:
                  - /placeholder: Find in this sheet…
                - button "Find previous" [ref=e93] [cursor=pointer]
                - button "Find next" [ref=e96] [cursor=pointer]
              - generic [ref=e100]:
                - combobox "Sheet actions" [ref=e101]:
                  - option "Sheet actions" [selected]
                  - option "Sort selection A → Z"
                  - option "Sort selection Z → A"
                  - option "Custom sort…"
                  - option "Paste special…"
                  - option "Fill series…"
                  - option "Cell note and link…"
                  - option "Freeze rows and columns before selection"
                  - option "Freeze first row"
                  - option "Freeze first column"
                  - option "Unfreeze rows and columns"
                  - option "Set selected row height…"
                  - option "Set selected column width…"
                  - option "Insert rows above selection"
                  - option "Insert columns before selection"
                  - option "Delete selected rows"
                  - option "Delete selected columns"
                  - option "Hide selected rows"
                  - option "Hide selected columns"
                  - option "Show selected rows"
                  - option "Show selected columns"
                  - option "Show all rows"
                  - option "Show all columns"
                  - option "Filter column by search"
                  - option "Custom filter…"
                  - option "Clear filter"
                  - option "Rename sheet"
                  - option "Delete sheet"
                  - option "Replace literal text…"
                - generic [ref=e102]:
                  - checkbox "Read only" [ref=e103]
                  - text: Read only
                - button "Table view" [ref=e104] [cursor=pointer]
                - combobox "Zoom" [ref=e105]:
                  - option "100%" [selected]
                  - option "75%"
                  - option "125%"
                  - option "150%"
                  - option "200%"
            - generic [ref=e107]:
              - toolbar "Spreadsheet tools" [ref=e108]:
                - generic [ref=e109]:
                  - generic [ref=e110]:
                    - button "Undo" [ref=e111] [cursor=pointer]
                    - button "Redo" [ref=e115] [cursor=pointer]
                  - combobox "Number format" [ref=e120]:
                    - option "Automatic" [selected]
                    - option "Number"
                    - option "Currency"
                    - option "Percent"
                    - option "Date"
                  - generic [ref=e121]:
                    - button "Bold" [ref=e122] [cursor=pointer]
                    - button "Italic" [ref=e125] [cursor=pointer]
                    - button "Underline" [ref=e128] [cursor=pointer]
                  - generic [ref=e131]:
                    - textbox "Cell background" [ref=e132] [cursor=pointer]: "#e1efe7"
                    - button "Align left" [ref=e133] [cursor=pointer]
                    - button "Align center" [ref=e135] [cursor=pointer]
                    - button "Align right" [ref=e137] [cursor=pointer]
                - button "More tools" [ref=e140] [cursor=pointer]
              - generic [ref=e145]:
                - textbox "Cell address" [ref=e146]: A1
                - generic [ref=e147]: ƒx
                - combobox "Formula bar" [ref=e149]: Item
              - grid "Spreadsheet" [ref=e151]:
                - row [ref=e154]:
                  - 'gridcell "A1: Item" [ref=e155]'
                - scrollbar "Horizontal scroll" [ref=e156]
                - scrollbar "Vertical scroll" [ref=e157]
              - generic [ref=e158]:
                - tablist "Sheets" [ref=e159]:
                  - tab "Launch budget" [selected] [ref=e160] [cursor=pointer]
                  - tab "Notes" [ref=e161] [cursor=pointer]
                  - button "Add sheet" [ref=e162] [cursor=pointer]
                - generic [ref=e164]: Ready · Local only
          - separator "Resize spreadsheet and results" [ref=e165]
          - complementary "Demo results" [ref=e167]:
            - tablist "Result panel" [ref=e168]:
              - tab "Data view" [selected] [ref=e169] [cursor=pointer]
              - tab "Code export" [ref=e170] [cursor=pointer]
            - tabpanel "Data view" [ref=e171]:
              - generic [ref=e172]:
                - generic [ref=e173]:
                  - generic [ref=e174]:
                    - generic [ref=e175]: Planned spend
                    - strong [ref=e176]: $13,097
                    - generic [ref=e177]: Calculated from the workbook
                  - generic [ref=e178]:
                    - generic [ref=e179]: Launch reserve
                    - strong [ref=e180]: $4,903
                    - generic [ref=e181]: Against an $18,000 demo budget
                  - generic [ref=e182]:
                    - generic [ref=e183]: Budget allocated
                    - strong [ref=e184]: 73%
                    - generic [ref=e185]: Edit quantity or unit cost below
                - img "Budget allocated 73 percent" [ref=e186]
              - paragraph [ref=e188]: Edit the sheet. Your changes update this view in real time.
```

# Test source

```ts
  1  | import { expect, test } from '@playwright/test';
  2  | 
  3  | for (const cacheBytes of [0, 4 * 1024 * 1024]) {
  4  |   test(`overlapping formulas preserve results and history with cache ${cacheBytes}`, async ({
  5  |     page,
  6  |   }) => {
  7  |     await page.goto('/');
  8  |     await expect(page.locator('#scene-select')).toBeEnabled();
> 9  |     const result = await page.evaluate(
     |                               ^ Error: page.evaluate: Test timeout of 30000ms exceeded.
  10 |       async ({ moduleUrl, cacheBytes }) => {
  11 |         const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ moduleUrl);
  12 |         let book = await createWorkbook(
  13 |           {
  14 |             sheets: [
  15 |               { name: 'Source', rows: 512, columns: 2 },
  16 |               { name: 'Output', rows: 512, columns: 2 },
  17 |             ],
  18 |           },
  19 |           { cacheBytes },
  20 |         );
  21 |         try {
  22 |           const [source, output] = book.getSheets();
  23 |           await source.range('A1:A32').setValues(Array.from({ length: 32 }, () => [1]));
  24 |           await output
  25 |             .range('A1:A512')
  26 |             .setFormulas(Array.from({ length: 512 }, (_, i) => [`SUM(Source!$A$1:$A$32)+${i}`]));
  27 |           const verify = async (expectedSource: number) => {
  28 |             const values = await book.getSheets()[1].range('A1:A512').getValues();
  29 |             for (let i = 0; i < values.length; i++)
  30 |               if (values[i][0] !== expectedSource + 31 + i)
  31 |                 throw new Error(`Wrong result at ${i}: ${JSON.stringify(values[i])}`);
  32 |             return values.length;
  33 |           };
  34 |           const counts = [await verify(1)];
  35 |           await source.range('A1').setValues([[2]]);
  36 |           counts.push(await verify(2));
  37 |           const undo = await book.undo();
  38 |           counts.push(await verify(1));
  39 |           const redo = await book.redo();
  40 |           counts.push(await verify(2));
  41 |           const id = book.id;
  42 |           await book.close();
  43 |           book = await openWorkbook(id, { cacheBytes });
  44 |           counts.push(await verify(2));
  45 |           return { counts, undo, redo };
  46 |         } finally {
  47 |           await book.deleteStorage();
  48 |           book.dispose();
  49 |         }
  50 |       },
  51 |       { moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`, cacheBytes },
  52 |     );
  53 |     expect(result).toEqual({ counts: [512, 512, 512, 512, 512], undo: true, redo: true });
  54 |   });
  55 | }
  56 | 
```
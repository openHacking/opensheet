# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: startup-recovery.spec.ts >> opening a saved scene does not overwrite a newer scene reference
- Location: tests/e2e/startup-recovery.spec.ts:3:1

# Error details

```
Error: expect(locator).toHaveText(expected) failed

Locator:  locator('#scene-title')
Expected: "Budget Simulator"
Received: ""
Timeout:  5000ms

Call log:
  - Expect "toHaveText" locator('#scene-title') with timeout 5000ms
  - waiting for locator('#scene-title')
    5 × locator resolved to <h1 id="scene-title"></h1>
      - unexpected value ""

```

```yaml
- complementary:
  - link "OpenSheet.":
    - /url: "#budget"
  - text: O Open workspace Open-source workspace WORKSPACE
  - button "Spreadsheet 1"
  - text: DEVELOPER TOOLS
  - button "API quick start"
  - button "Plugin example"
  - text: Your data stays yours.
  - paragraph: Files are processed on your device. No account. No uploads.
  - text: Open source v0.2.0-dev.0
- banner:
  - text: Workspace / Playground Runs locally
  - button "Keyboard shortcuts"
  - text: OS
- main:
  - region "Demo controls":
    - text: DEMO
    - combobox "Choose demo" [disabled]:
      - option "Performance Lab" [selected]
      - option "Budget Simulator"
      - option "Sales Dashboard"
      - option "Project Planner"
      - option "Table to Code"
    - heading [level=1]
    - paragraph
    - button "Reset demo"
  - region "Spreadsheet editor":
    - textbox "Workbook name": Launch budget
    - status: Opening…
    - button "Import file"
    - button "Save"
    - button "Export JSON"
    - button "Export XLSX"
    - textbox "Find in sheet":
      - /placeholder: Find in this sheet…
    - button "Find previous"
    - button "Find next"
    - combobox "Sheet actions":
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
    - checkbox "Read only"
    - text: Read only
    - button "Table view"
    - combobox "Zoom":
      - option "100%" [selected]
      - option "75%"
      - option "125%"
      - option "150%"
      - option "200%"
    - toolbar "Spreadsheet tools":
      - button "Undo"
      - button "Redo"
      - combobox "Number format":
        - option "Automatic" [selected]
        - option "Number"
        - option "Currency"
        - option "Percent"
        - option "Date"
      - button "Bold"
      - button "Italic"
      - button "Underline"
      - textbox "Cell background": "#e1efe7"
      - button "Align left"
      - button "Align center"
      - button "Align right"
      - button "More tools"
    - textbox "Cell address": A1
    - text: ƒx
    - combobox "Formula bar"
    - tablist "Sheets"
    - text: Ready
  - separator "Resize spreadsheet and results"
  - complementary "Demo results":
    - tablist "Result panel":
      - tab "Data view" [selected]
      - tab "Code export"
    - tabpanel "Data view":
      - paragraph: Edit the sheet. Your changes update this view in real time.
```

# Test source

```ts
  1   | import { expect, test } from '@playwright/test';
  2   | 
  3   | test('opening a saved scene does not overwrite a newer scene reference', async ({
  4   |   page,
  5   |   context,
  6   | }) => {
  7   |   await page.goto('/#budget');
  8   |   await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  9   |   const oldId = await page.evaluate(async () => {
  10  |     const book = window.opensheet.getWorkbook();
  11  |     await book
  12  |       .getSheets()[0]
  13  |       .range('A1')
  14  |       .setValues([[314]]);
  15  |     return book.id;
  16  |   });
  17  |   page.once('dialog', (dialog) => dialog.accept());
  18  |   await page.locator('#scene-select').selectOption('sales');
  19  |   await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  20  |   const newId = await page.evaluate(async () => {
  21  |     const book = window.opensheet.getWorkbook();
  22  |     await book
  23  |       .getSheets()[0]
  24  |       .range('A1')
  25  |       .setValues([[271]]);
  26  |     await book.close();
  27  |     return book.id;
  28  |   });
  29  |   const opening = await context.newPage();
  30  |   await opening.addInitScript(
  31  |     ({ next }) => {
  32  |       if (sessionStorage.getItem('reference-race-injected')) return;
  33  |       const get = Storage.prototype.getItem;
  34  |       Storage.prototype.getItem = function (key) {
  35  |         const value = get.call(this, key);
  36  |         if (this === localStorage && key === 'opensheet.scene.budget') {
  37  |           Storage.prototype.getItem = get;
  38  |           sessionStorage.setItem('reference-race-injected', 'yes');
  39  |           localStorage.setItem(key, next);
  40  |         }
  41  |         return value;
  42  |       };
  43  |     },
  44  |     { next: newId },
  45  |   );
  46  |   await opening.goto('/#budget');
> 47  |   await expect(opening.locator('#scene-title')).toHaveText('Budget Simulator');
      |                                                 ^ Error: expect(locator).toHaveText(expected) failed
  48  |   expect(await opening.evaluate(() => window.opensheet.getWorkbook().id)).toBe(oldId);
  49  |   expect(await opening.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(newId);
  50  |   await opening.reload();
  51  |   await expect(opening.locator('#scene-title')).toHaveText('Budget Simulator');
  52  |   expect(await opening.evaluate(() => window.opensheet.getWorkbook().id)).toBe(newId);
  53  |   expect(
  54  |     await opening.evaluate(async () =>
  55  |       window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
  56  |     ),
  57  |   ).toEqual([[271]]);
  58  | });
  59  | 
  60  | test('failed restore preserves the saved reference and can retry after storage is repaired', async ({
  61  |   page,
  62  |   context,
  63  | }) => {
  64  |   await page.goto('/#budget');
  65  |   await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  66  |   const id = await page.evaluate(async () => {
  67  |     const book = window.opensheet.getWorkbook();
  68  |     await book
  69  |       .getSheets()[0]
  70  |       .range('A1')
  71  |       .setValues([[314]]);
  72  |     await book.close();
  73  |     const open = indexedDB.open('opensheet');
  74  |     const db = await new Promise<IDBDatabase>((resolve, reject) => {
  75  |       open.onsuccess = () => resolve(open.result);
  76  |       open.onerror = () => reject(open.error);
  77  |     });
  78  |     const read = db.transaction('heads').objectStore('heads').get(book.id);
  79  |     (window as any).savedHead = await new Promise((resolve, reject) => {
  80  |       read.onsuccess = () => resolve(read.result);
  81  |       read.onerror = () => reject(read.error);
  82  |     });
  83  |     const tx = db.transaction('heads', 'readwrite');
  84  |     tx.objectStore('heads').delete(book.id);
  85  |     await new Promise<void>((resolve, reject) => {
  86  |       tx.oncomplete = () => resolve();
  87  |       tx.onabort = () => reject(tx.error);
  88  |     });
  89  |     db.close();
  90  |     return book.id;
  91  |   });
  92  |   const retryPage = await context.newPage();
  93  |   await retryPage.goto('/#budget');
  94  |   await expect(retryPage.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  95  |   expect(await retryPage.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
  96  |   await expect(retryPage.locator('#app')).toHaveAttribute('inert', '');
  97  |   await page.evaluate(async (id) => {
  98  |     const open = indexedDB.open('opensheet');
  99  |     const db = await new Promise<IDBDatabase>((resolve, reject) => {
  100 |       open.onsuccess = () => resolve(open.result);
  101 |       open.onerror = () => reject(open.error);
  102 |     });
  103 |     const tx = db.transaction('heads', 'readwrite');
  104 |     tx.objectStore('heads').put((window as any).savedHead, id);
  105 |     await new Promise<void>((resolve, reject) => {
  106 |       tx.oncomplete = () => resolve();
  107 |       tx.onabort = () => reject(tx.error);
  108 |     });
  109 |     db.close();
  110 |   }, id);
  111 |   await retryPage.getByRole('button', { name: 'Try again' }).click();
  112 |   await expect(retryPage.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  113 |   await expect(retryPage.locator('#scene-title')).toHaveText('Budget Simulator');
  114 |   expect(
  115 |     await retryPage.evaluate(
  116 |       async () =>
  117 |         (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
  118 |     ),
  119 |   ).toBe(314);
  120 |   expect(await retryPage.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
  121 | });
  122 | 
  123 | test('unavailable browser storage reports a recoverable startup error', async ({ page }) => {
  124 |   await page.addInitScript(() => {
  125 |     const original = Object.getOwnPropertyDescriptor(window, 'indexedDB');
  126 |     Object.defineProperty(window, 'indexedDB', { configurable: true, value: undefined });
  127 |     (window as any).restoreIndexedDB = () =>
  128 |       original
  129 |         ? Object.defineProperty(window, 'indexedDB', original)
  130 |         : delete (window as any).indexedDB;
  131 |   });
  132 |   await page.goto('/#budget');
  133 |   await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  134 |   await expect(page.locator('#startup-recovery-message')).toContainText('STORAGE_UNAVAILABLE');
  135 |   await expect(page.getByRole('button', { name: 'Try again' })).toBeFocused();
  136 |   await page.evaluate(() => (window as any).restoreIndexedDB());
  137 |   await page.getByRole('button', { name: 'Try again' }).click();
  138 |   await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  139 |   await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  140 | });
  141 | 
  142 | test('denied browser storage access reports guidance and retries after access is restored', async ({
  143 |   page,
  144 | }) => {
  145 |   await page.addInitScript(() => {
  146 |     const original = Object.getOwnPropertyDescriptor(window, 'indexedDB');
  147 |     Object.defineProperty(window, 'indexedDB', {
```
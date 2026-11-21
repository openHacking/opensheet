import { expect, test } from '@playwright/test';

for (const format of ['json', 'binary'] as const) {
  for (const change of ['revision', 'reference', 'draft'] as const) {
    test(`${format} backup restore rejects a ${change} change during preparation`, async ({
      page,
      context,
    }) => {
      await page.goto('/#budget');
      await expect(page.locator('#scene-select')).toBeEnabled();
      const fixture = await page.evaluate(async () => {
        const book = window.opensheet.getWorkbook();
        await book
          .getSheets()[0]
          .range('A1')
          .setValues([[9876]]);
        const parts: string[] = [];
        for await (const part of book.streamJSON()) parts.push(part);
        const binary: number[] = [];
        for await (const part of book.streamBinary()) binary.push(...part);
        await book.close();
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const tx = db.transaction('heads', 'readwrite');
        tx.objectStore('heads').delete(book.id);
        await new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve();
          tx.onabort = () => reject(tx.error);
        });
        db.close();
        return { id: book.id, json: parts.join(''), binary };
      });
      const recovery = await context.newPage();
      await recovery.goto('/#budget');
      await expect(
        recovery.getByRole('heading', { name: 'Unable to open workbook' }),
      ).toBeVisible();
      const previousId = await recovery.evaluate(
        async ({ json, format }) => {
          const app = window.opensheet;
          await app.load(JSON.parse(json));
          const previous = app.getWorkbook();
          const prototype = Object.getPrototypeOf(previous);
          const method = format === 'json' ? 'importJSON' : 'importBinary';
          const original = prototype[method];
          prototype[method] = async function (...args: unknown[]) {
            const result = await original.apply(this, args);
            if (this !== previous) {
              prototype[method] = original;
              await new Promise<void>((resolve) => ((window as any).resumeBackupImport = resolve));
            }
            return result;
          };
          return previous.id;
        },
        { json: fixture.json, format },
      );
      const heads = () =>
        recovery.evaluate(async () => {
          const request = indexedDB.open('opensheet');
          const db = await new Promise<IDBDatabase>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          });
          const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
            const read = db.transaction('heads').objectStore('heads').getAllKeys();
            read.onsuccess = () => resolve(read.result);
            read.onerror = () => reject(read.error);
          });
          db.close();
          return keys.map(String).sort();
        });
      const before = await heads();
      const input = recovery.locator('#restore-backup-file');
      const backup = {
        name: format === 'json' ? 'backup.json' : 'backup.opensheet',
        mimeType: 'application/octet-stream',
        buffer: format === 'json' ? Buffer.from(fixture.json) : Buffer.from(fixture.binary),
      };
      await input.setInputFiles(backup);
      await expect
        .poll(() => recovery.evaluate(() => !!(window as any).resumeBackupImport))
        .toBe(true);
      if (change === 'revision') {
        await recovery.evaluate(() =>
          window.opensheet
            .getWorkbook()
            .getSheets()[0]
            .range('A1')
            .setValues([[4321]]),
        );
      } else if (change === 'reference') {
        await page.evaluate(() =>
          localStorage.setItem('opensheet.scene.budget', 'other-reference'),
        );
        await expect
          .poll(() => recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget')))
          .toBe('other-reference');
      } else {
        await recovery.evaluate(() => {
          document.getElementById('app')!.inert = false;
        });
        await recovery.getByRole('combobox', { name: 'Formula bar' }).fill('unsaved draft');
        expect(await recovery.evaluate(() => window.opensheet.hasDraft)).toBe(true);
      }
      await recovery.evaluate(() => (window as any).resumeBackupImport());
      await expect(recovery.locator('#startup-recovery-message')).toContainText(
        'changed during backup restore',
      );
      await expect(recovery.getByRole('button', { name: 'Restore backup' })).toBeEnabled();
      expect(await heads()).toEqual(before);
      expect(
        await recovery.evaluate(() => ({
          id: window.opensheet.getWorkbook().id,
          reference: localStorage.getItem('opensheet.scene.budget'),
        })),
      ).toEqual({
        id: previousId,
        reference: change === 'reference' ? 'other-reference' : fixture.id,
      });
      if (change === 'revision')
        expect(
          await recovery.evaluate(() =>
            window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
          ),
        ).toEqual([[4321]]);
      if (change === 'draft') {
        const formula = recovery.getByRole('combobox', { name: 'Formula bar' });
        await expect(formula).toHaveValue('unsaved draft');
        await formula.press('Escape');
      }
      if (change === 'reference')
        await page.evaluate((id) => localStorage.setItem('opensheet.scene.budget', id), fixture.id);
      await input.setInputFiles(backup);
      await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
      expect(
        await recovery.evaluate(() =>
          window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
        ),
      ).toEqual([[9876]]);
      await recovery.reload();
      await expect(recovery.locator('#scene-select')).toBeEnabled();
      expect(
        await recovery.evaluate(() =>
          window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
        ),
      ).toEqual([[9876]]);
    });
  }
}

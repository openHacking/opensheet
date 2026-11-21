import { expect, test } from '@playwright/test';

for (const operation of ['initialize', 'compare'] as const) {
  for (const phase of ['before-commit', 'after-commit'] as const) {
    test(`scene reference ${operation} recovers when a page closes ${phase}`, async ({
      page,
      context,
    }) => {
      await page.goto('/#budget');
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      const key = `termination-${operation}-${phase}-${Date.now()}`;
      const initial = await page.evaluate(
        async ({ key, operation }) => {
          const path = '/src/scene-references.ts';
          const api = await import(/* @vite-ignore */ path);
          return operation === 'compare' ? api.initializeSceneReference(key, 'old-workbook') : null;
        },
        { key, operation },
      );
      const writer = await context.newPage();
      await writer.goto('/#budget');
      await expect(writer.locator('#scene-title')).toHaveText('Budget Simulator');
      await writer.evaluate(
        async ({ key, initial, phase, operation }) => {
          const path = '/src/scene-references.ts';
          const api = await import(/* @vite-ignore */ path);
          if (phase === 'before-commit') {
            const put = IDBObjectStore.prototype.put;
            IDBObjectStore.prototype.put = function (...args) {
              const request = put.apply(this, args);
              if (this.name === 'references') {
                const store = this;
                request.addEventListener(
                  'success',
                  () => {
                    (window as any).referenceTerminationPhase = 'before-commit';
                    const keepAlive = () => {
                      const next = store.get(key);
                      next.onsuccess = keepAlive;
                    };
                    keepAlive();
                  },
                  { once: true },
                );
              }
              return request;
            };
          } else {
            const transaction = IDBDatabase.prototype.transaction;
            const oncomplete = Object.getOwnPropertyDescriptor(
              IDBTransaction.prototype,
              'oncomplete',
            )!;
            IDBDatabase.prototype.transaction = function (...args) {
              const tx = transaction.apply(this, args);
              if (this.name === 'opensheet-playground-references' && args[1] === 'readwrite') {
                Object.defineProperty(tx, 'oncomplete', {
                  configurable: true,
                  set(handler) {
                    oncomplete.set!.call(tx, (event: Event) => {
                      (window as any).referenceTerminationPhase = 'after-commit';
                      // Keep the helper response pending until this page is closed.
                      (window as any).releaseReferenceResponse = () => handler.call(tx, event);
                    });
                  },
                });
              }
              return tx;
            };
          }
          (window as any).referenceSettled = false;
          const pending =
            operation === 'compare'
              ? api.compareSceneReference(key, initial, 'new-workbook')
              : api.initializeSceneReference(key, 'legacy-workbook');
          void pending.then(
            () => ((window as any).referenceSettled = true),
            () => ((window as any).referenceSettled = true),
          );
        },
        { key, initial, phase, operation },
      );
      await expect
        .poll(() => writer.evaluate(() => (window as any).referenceTerminationPhase))
        .toBe(phase);
      expect(await writer.evaluate(() => (window as any).referenceSettled)).toBe(false);
      await writer.close();
      const after = await page.evaluate(
        async ({ key, initial, operation }) => {
          const path = '/src/scene-references.ts';
          const api = await import(/* @vite-ignore */ path);
          const recovered = await api.readSceneReference(key);
          const retry =
            operation === 'compare'
              ? await api.compareSceneReference(key, initial, 'retry-workbook')
              : await api.initializeSceneReference(key, 'retry-workbook');
          return { recovered, retry, final: await api.readSceneReference(key) };
        },
        { key, initial, operation },
      );
      if (operation === 'initialize' && phase === 'before-commit') {
        expect(after).toEqual({
          recovered: null,
          retry: { value: 'retry-workbook', revision: 0 },
          final: { value: 'retry-workbook', revision: 0 },
        });
      } else if (operation === 'initialize') {
        expect(after).toEqual({
          recovered: { value: 'legacy-workbook', revision: 0 },
          retry: { value: 'legacy-workbook', revision: 0 },
          final: { value: 'legacy-workbook', revision: 0 },
        });
      } else if (phase === 'before-commit') {
        expect(after).toEqual({
          recovered: { value: 'old-workbook', revision: 0 },
          retry: { updated: true, current: { value: 'retry-workbook', revision: 1 } },
          final: { value: 'retry-workbook', revision: 1 },
        });
      } else {
        expect(after).toEqual({
          recovered: { value: 'new-workbook', revision: 1 },
          retry: { updated: false, current: { value: 'new-workbook', revision: 1 } },
          final: { value: 'new-workbook', revision: 1 },
        });
      }
    });
  }
}

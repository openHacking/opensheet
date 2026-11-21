import { createOpenSheet, openWorkbook } from 'opensheet';
import 'opensheet/style.css';

// Export a portable copy separately from browser storage.
async function backupWorkbook(book) {
  return JSON.stringify(await book.toJSON());
}

async function restoreBackup(app, json) {
  await app.load(JSON.parse(json));
  return app.getWorkbook();
}

const storageKey = 'opensheet-local-candidate-workbook-id';
const maxBackupBytes = 64 * 1024 * 1024;
let app;
try {
  const savedId = localStorage.getItem(storageKey);
  app = createOpenSheet({ container: '#sheet' });
  let book;
  if (savedId) {
    await app.open(savedId);
    book = app.getWorkbook();
  } else {
    book = await app.createWorkbook({ sheets: [{ name: 'Example', rows: 10, columns: 4 }] });
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[7, 5]]);
    await sheet.range('C1').setFormulas([['A1+B1']]);
    localStorage.setItem(storageKey, book.id);
  }
  await app.ready();
  window.candidateExample = {
    app,
    book,
    openWorkbook,
    backupWorkbook,
    restoreBackup,
    createOpenSheet,
  };
  const backup = document.querySelector('#backup');
  const restore = document.querySelector('#restore');
  const status = document.querySelector('#status');
  const busy = (value) => {
    backup.disabled = value;
    restore.disabled = value;
  };
  busy(false);
  backup.addEventListener('click', async () => {
    busy(true);
    status.textContent = '';
    try {
      const json = await backupWorkbook(book);
      const blob = new Blob([json], { type: 'application/json' });
      if (blob.size > maxBackupBytes) throw new Error('Backup exceeds the 64 MiB example limit');
      const url = URL.createObjectURL(blob);
      try {
        const link = document.createElement('a');
        link.href = url;
        link.download = 'workbook.opensheet.json';
        link.click();
      } finally {
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }
    } catch (error) {
      status.textContent = `Could not create backup: ${error?.message ?? error}`;
    } finally {
      busy(false);
    }
  });
  restore.addEventListener('change', async () => {
    const file = restore.files?.[0];
    if (!file) return;
    busy(true);
    status.textContent = '';
    let restored = false;
    try {
      if (file.size > maxBackupBytes) throw new Error('Backup exceeds the 64 MiB example limit');
      book = await restoreBackup(app, await file.text());
      window.candidateExample.book = book;
      restored = true;
      localStorage.setItem(storageKey, book.id);
      status.textContent = 'Backup restored.';
    } catch (error) {
      status.textContent = restored
        ? `Workbook restored, but its ID could not be saved. Export a backup before refreshing: ${error?.message ?? error}`
        : `Could not restore backup; check the workbook before retrying: ${error?.message ?? error}`;
    } finally {
      restore.value = '';
      busy(false);
    }
  });
} catch (error) {
  try {
    app?.dispose();
  } catch {
    // Preserve the startup error.
  }
  document.querySelector('#status').textContent =
    `Could not open workbook: ${error?.message ?? error}`;
  document.body.dataset.error = String(error?.stack ?? error);
}

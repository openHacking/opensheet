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

try {
  const app = createOpenSheet({ container: '#sheet' });
  const book = await app.createWorkbook({ sheets: [{ name: 'Example', rows: 10, columns: 4 }] });
  const sheet = book.getSheets()[0];
  await sheet.range('A1:B1').setValues([[7, 5]]);
  await sheet.range('C1').setFormulas([['A1+B1']]);
  await app.ready();
  window.candidateExample = {
    app,
    book,
    openWorkbook,
    backupWorkbook,
    restoreBackup,
    createOpenSheet,
  };
} catch (error) {
  document.body.dataset.error = String(error?.stack ?? error);
}

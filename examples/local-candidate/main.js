import { createOpenSheet, openWorkbook } from 'opensheet';
import 'opensheet/style.css';

try {
  const app = createOpenSheet({ container: '#sheet' });
  const book = await app.createWorkbook({ sheets: [{ name: 'Example', rows: 10, columns: 4 }] });
  const sheet = book.getSheets()[0];
  await sheet.range('A1:B1').setValues([[7, 5]]);
  await sheet.range('C1').setFormulas([['A1+B1']]);
  await app.ready();
  window.candidateExample = { app, book, openWorkbook };
} catch (error) {
  document.body.dataset.error = String(error?.stack ?? error);
}

import { readFile } from 'node:fs/promises';
import { FileReader, parseRange, validateWorkbookFile } from '@opensheetjs/core';

const [path, sheetName, address] = process.argv.slice(2);
if (!path || !sheetName || !address || process.argv.length !== 5) {
  console.error('Usage: node node-file-reader.mjs FILE.json SHEET_NAME A1_RANGE');
  process.exitCode = 2;
} else {
  const bounds = parseRange(address);
  const size = (bounds.endRow - bounds.startRow) * (bounds.endColumn - bounds.startColumn);
  if (size > 10000) throw new Error('Example reads at most 10,000 cells at a time');
  const file = validateWorkbookFile(JSON.parse(await readFile(path, 'utf8')));
  const sheet = file.sheets.find((item) => item.name.toLowerCase() === sheetName.toLowerCase());
  if (!sheet) throw new Error(`Unknown sheet: ${sheetName}`);
  const reader = new FileReader(file);
  try {
    const values = reader.getSheetById(sheet.id).range(bounds).getValues();
    console.log(JSON.stringify({ sheet: sheet.name, range: address, values }));
  } finally {
    reader.dispose();
  }
}

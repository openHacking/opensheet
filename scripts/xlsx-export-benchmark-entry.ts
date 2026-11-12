import * as XLSX from 'xlsx';
import { writeXlsx } from '@opensheetjs/adapter-sheetjs';
import type { WorkbookFile } from '@opensheetjs/core';

export function write(file: WorkbookFile) {
  return writeXlsx(file, XLSX);
}

export function compare(first: Uint8Array, second: Uint8Array) {
  const a = XLSX.read(first, { type: 'array', cellFormula: true, bookFiles: true });
  const b = XLSX.read(second, { type: 'array', cellFormula: true, bookFiles: true });
  if (JSON.stringify(a.SheetNames) !== JSON.stringify(b.SheetNames))
    throw new Error('Sheet names differ');
  let cells = 0;
  for (const name of a.SheetNames) {
    const left = a.Sheets[name],
      right = b.Sheets[name];
    const keys = Object.keys(left)
      .filter((key) => !key.startsWith('!'))
      .sort();
    const other = Object.keys(right)
      .filter((key) => !key.startsWith('!'))
      .sort();
    if (keys.length !== other.length) throw new Error('Cell counts differ');
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i],
        x = left[key],
        y = right[key];
      if (key !== other[i] || x.t !== y.t || x.v !== y.v || x.f !== y.f)
        throw new Error(`Cell differs: ${name}!${key}`);
      cells++;
    }
  }
  const path = 'xl/worksheets/sheet1.xml';
  const firstXml = a.files?.[path]?.content,
    secondXml = b.files?.[path]?.content;
  const xml = firstXml && new TextDecoder().decode(firstXml);
  if (!xml || !secondXml || xml !== new TextDecoder().decode(secondXml))
    throw new Error('Worksheet XML differs');
  return { cells, sheets: a.SheetNames.length, frozen: xml.includes('state="frozen"') };
}

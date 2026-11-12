import * as XLSX from 'xlsx';

/** A normal XLSX with the first underscore represented by an XML character reference. */
export function xlsxEntityEscape(): Buffer {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['_x0041_']]), 'Data');
  const bytes = XLSX.write(book, { type: 'buffer', bookType: 'xlsx' });
  const archive = XLSX.CFB.read(bytes, { type: 'buffer' });
  const entry = XLSX.CFB.find(archive, '/xl/worksheets/sheet1.xml');
  const xml = new TextDecoder().decode(entry.content);
  if (!xml.includes('_x0041_')) throw new Error('Fixture has no target text');
  XLSX.CFB.utils.cfb_add(
    archive,
    '/xl/worksheets/sheet1.xml',
    Buffer.from(xml.replace('_x0041_', '&#95;x0041_')),
  );
  return XLSX.CFB.write(archive, { type: 'buffer', fileType: 'zip' });
}

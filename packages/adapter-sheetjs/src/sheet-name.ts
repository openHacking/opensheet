/** Sheet-name restrictions enforced by the pinned SheetJS XLSX writer. */
export function unsupportedXlsxSheetName(name: string): boolean {
  return (
    typeof name !== 'string' ||
    name.length === 0 ||
    name.length > 31 ||
    name.startsWith("'") ||
    name.endsWith("'") ||
    name.toLowerCase() === 'history' ||
    /[\\/?*\[\]:]/.test(name)
  );
}

export const XLSX_SHEET_NAME_MESSAGE =
  'Rename sheets named History or beginning or ending with an apostrophe before exporting XLSX; names must also be 1–31 characters and exclude : \\ / ? * [ ]';

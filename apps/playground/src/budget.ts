import { validateWorkbookFile, type WorkbookFile } from 'opensheet';
import fixture from './budget-fixture.js';

/** A fresh native snapshot avoids creating a temporary persisted workbook at startup. */
export async function sample(): Promise<WorkbookFile> {
  return validateWorkbookFile(structuredClone(fixture));
}

import type { Page } from '@playwright/test';

/** A read racing a commit is retried by expect.poll, while unrelated errors still fail. */
export function readForPoll(page: Page, address: string) {
  return page.evaluate(async (address) => {
    try {
      return await window.opensheet.getWorkbook().getSheets()[0].range(address).getValues();
    } catch (error) {
      if ((error as { code?: string }).code === 'REVISION_CONFLICT') return 'retry';
      throw error;
    }
  }, address);
}

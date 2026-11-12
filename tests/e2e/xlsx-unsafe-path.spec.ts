import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import {
  byteCaseCollisionZip,
  storedDescriptorZip,
  twoEntryZip,
  withEndSignature,
} from '../zip-fixtures.js';

test('imports independently authored stored streaming XLSX and preserves data after reopen', async ({
  page,
}) => {
  const buffer = readFileSync(
    new URL('../../benchmarks/streaming-stored-openpyxl-2026-11-11-fixed.xlsx', import.meta.url),
  );
  expect(createHash('sha256').update(buffer).digest('hex')).toBe(
    'b8c001c4d942d4723fab7eb84dfde3ac01e538618bcaff3e14285c65482ac4f2',
  );
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'independent-stream.xlsx',
    mimeType: 'application/octet-stream',
    buffer,
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Data validation rules are not imported');
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Independent stream', exact: true })).toBeVisible();
  const state = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const first = await sheet.range('A1:B2').getValues();
    const formula = await sheet.range('B1').getFormulas();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { first, formula, reopened: await sheet.range('A1:B2').getValues() };
  });
  expect(state.first).toEqual([
    [42, 84],
    ['streamed text', null],
  ]);
  expect(state.formula).toEqual([['A1*2']]);
  expect(state.reopened).toEqual(state.first);
});

for (const signed of [true, false])
  test(`imports a verified stored entry (signed=${signed})`, async ({ page }) => {
    const workbook = XLSX.utils.book_new(XLSX.utils.aoa_to_sheet([[42]]), 'Descriptor data');
    const buffer = storedDescriptorZip(
      XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', compression: false }),
      signed,
    );
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    await page
      .locator('#file')
      .setInputFiles({ name: 'descriptor.xlsx', mimeType: 'application/octet-stream', buffer });
    const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
    await expect(dialog).toBeVisible();
    await dialog.getByRole('button', { name: 'Close dialog' }).click();
    await expect(page.getByRole('tab', { name: 'Descriptor data', exact: true })).toBeVisible();
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[42]]);
  });

for (const [name, buffer, message] of [
  ['legacy-byte-collision', byteCaseCollisionZip(), 'duplicate entry name'],
  ['end-record-collision', withEndSignature(twoEntryZip(['a', 'b'])), 'ambiguous end record'],
] as const) {
  test(`rejects ${name} before replacing the workbook`, async ({ page }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    }));
    await page
      .locator('#file')
      .setInputFiles({ name: `${name}.xlsx`, mimeType: 'application/octet-stream', buffer });
    await expect(page.locator('#notice')).toContainText(message);
    expect(
      await page.evaluate(() => ({
        id: window.opensheet.getWorkbook().id,
        revision: window.opensheet.getWorkbook().revision,
      })),
    ).toEqual(before);
  });
}

test('rejects case-colliding ZIP parts without replacing the workbook', async ({ page }) => {
  const buffer = twoEntryZip(['xl/sheet.xml', 'XL/SHEET.XML']);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => ({
    id: window.opensheet.getWorkbook().id,
    revision: window.opensheet.getWorkbook().revision,
  }));
  await page
    .locator('#file')
    .setInputFiles({ name: 'ambiguous.xlsx', mimeType: 'application/octet-stream', buffer });
  await expect(page.locator('#notice')).toContainText('duplicate entry name');
  expect(
    await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    })),
  ).toEqual(before);
});

test('rejects truncated ZIP timestamp before replacing the workbook', async ({ page }) => {
  const name = Buffer.from('entry');
  const extra = Buffer.from([0x55, 0x54, 1, 0, 1]);
  const data = Buffer.from('abc');
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50);
  local.writeUInt16LE(20, 4);
  local.writeUInt32LE(0x352441c2, 14);
  local.writeUInt32LE(data.length, 18);
  local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(name.length, 26);
  local.writeUInt16LE(extra.length, 28);
  const directory = Buffer.alloc(46);
  directory.writeUInt32LE(0x02014b50);
  directory.writeUInt16LE(20, 6);
  directory.writeUInt32LE(0x352441c2, 16);
  directory.writeUInt32LE(data.length, 20);
  directory.writeUInt32LE(data.length, 24);
  directory.writeUInt16LE(name.length, 28);
  directory.writeUInt16LE(extra.length, 30);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(1, 8);
  end.writeUInt16LE(1, 10);
  end.writeUInt32LE(directory.length + name.length + extra.length, 12);
  end.writeUInt32LE(local.length + name.length + extra.length + data.length, 16);
  const bytes = Buffer.concat([local, name, extra, data, directory, name, extra, end]);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => ({
    id: window.opensheet.getWorkbook().id,
    revision: window.opensheet.getWorkbook().revision,
  }));
  await page
    .locator('#file')
    .setInputFiles({ name: 'timestamp.xlsx', mimeType: 'application/octet-stream', buffer: bytes });
  await expect(page.locator('#notice')).toContainText('Invalid ZIP: invalid extended timestamp');
  expect(
    await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    })),
  ).toEqual(before);
});

for (const fault of ['traversal', 'drive', 'invalid-utf8'] as const) {
  test(`rejects XLSX ${fault} entry without replacing the workbook`, async ({ page }) => {
    const workbook = XLSX.utils.book_new(XLSX.utils.aoa_to_sheet([[42]]), 'Data');
    const bytes = Buffer.from(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }));
    const end = bytes.length - 22;
    let position = bytes.readUInt32LE(end + 16);
    const count = bytes.readUInt16LE(end + 10);
    let changed = false;
    for (let index = 0; index < count; index++) {
      const nameLength = bytes.readUInt16LE(position + 28);
      const extraLength = bytes.readUInt16LE(position + 30);
      const commentLength = bytes.readUInt16LE(position + 32);
      if (
        bytes.subarray(position + 46, position + 46 + nameLength).toString() === 'xl/workbook.xml'
      ) {
        const local = bytes.readUInt32LE(position + 42);
        if (fault === 'invalid-utf8') {
          bytes.writeUInt16LE(bytes.readUInt16LE(position + 8) | 0x0800, position + 8);
          bytes.writeUInt16LE(bytes.readUInt16LE(local + 6) | 0x0800, local + 6);
          bytes[position + 46] = 0xff;
          bytes[local + 30] = 0xff;
        } else {
          const name = fault === 'drive' ? 'C:/workbook.xml' : '../workbook.xml';
          bytes.write(name, position + 46);
          bytes.write(name, local + 30);
        }
        changed = true;
        break;
      }
      position += 46 + nameLength + extraLength + commentLength;
    }
    expect(changed).toBe(true);
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    }));
    await page
      .locator('#file')
      .setInputFiles({ name: 'unsafe.xlsx', mimeType: 'application/octet-stream', buffer: bytes });
    await expect(page.locator('#notice')).toContainText(
      fault === 'invalid-utf8'
        ? 'Invalid ZIP: invalid UTF-8 entry name'
        : 'Invalid ZIP: unsafe entry path',
    );
    expect(
      await page.evaluate(() => ({
        id: window.opensheet.getWorkbook().id,
        revision: window.opensheet.getWorkbook().revision,
      })),
    ).toEqual(before);
  });
}

import { expect, it } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';

it('rejects an incomplete search instead of reporting a partial result', async () => {
  const measure = process.env.OPENSHEET_TEST_TIMING === '1';
  let last = performance.now();
  let lastCpu = measure ? process.cpuUsage() : undefined;
  const timings: Record<string, { wallMs: number; processCpuMs: number }> = {};
  const checkpoint = (name: string) => {
    if (!lastCpu) return;
    const now = performance.now();
    const cpu = process.cpuUsage();
    timings[name] = {
      wallMs: now - last,
      processCpuMs: (cpu.user - lastCpu.user + cpu.system - lastCpu.system) / 1000,
    };
    last = now;
    lastCpu = cpu;
  };
  let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 130, columns: 2 }] });
  checkpoint('create');
  try {
    let sheet = book.getSheets()[0];
    await sheet
      .range('A1:A129')
      .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 2}+1`]));
    await sheet.range('A130').setValues([[1]]);
    await sheet.range('B1').setValues([['#LIMIT!']]);
    checkpoint('populate');
    const id = book.id;
    await book.close();
    book = await openWorkbook(id, { cacheBytes: 0 });
    checkpoint('reopen');
    sheet = book.getSheets()[0];
    const profilePath = process.env.OPENSHEET_SEARCH_CPU_PROFILE;
    const session = profilePath
      ? new (await import('node:inspector/promises')).Session()
      : undefined;
    if (session) {
      session.connect();
      await session.post('Profiler.enable');
      await session.post('Profiler.start');
    }
    try {
      for (const direction of ['next', 'previous'] as const)
        await expect(book.search(sheet.id, '#LIMIT!', -1, { direction })).rejects.toMatchObject({
          code: 'LIMIT_EXCEEDED',
        });
    } finally {
      if (session && profilePath) {
        const { profile } = await session.post('Profiler.stop');
        session.disconnect();
        await (await import('node:fs/promises')).writeFile(profilePath, JSON.stringify(profile));
      }
    }
    checkpoint('limited-searches');
    await sheet.range('A1:A129').clear();
    expect(await book.search(sheet.id, '#LIMIT!')).toEqual({
      count: 1,
      match: { row: 0, column: 1 },
    });
    checkpoint('clear-and-retry');
  } finally {
    await book.deleteStorage();
    await book.close();
    checkpoint('cleanup');
    if (measure) console.info('search-limit timings (ms)', JSON.stringify(timings));
  }
});

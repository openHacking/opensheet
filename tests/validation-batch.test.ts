import { expect, it, vi } from 'vitest';
import fc from 'fast-check';
import {
  checkStagedValidation,
  checkValidationBatch,
} from '../packages/core/src/validation-batch.js';
import { validationRuleSchema } from '../packages/core/src/validation-rule.js';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';

const range = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
const reject = validationRuleSchema.parse({
  kind: 'number',
  operator: 'gte',
  first: 10,
  allowBlank: false,
  policy: 'reject',
});
const warn = validationRuleSchema.parse({
  kind: 'number',
  operator: 'lte',
  first: 5,
  allowBlank: false,
  policy: 'warn',
});

it('matches an independent model for overlapping and disjoint candidate positions', async () => {
  await fc.assert(
    fc.asyncProperty(
      fc.array(
        fc.record({
          r: fc.integer({ min: 0, max: 7 }),
          c: fc.integer({ min: 0, max: 7 }),
          h: fc.integer({ min: 1, max: 3 }),
          w: fc.integer({ min: 1, max: 3 }),
          threshold: fc.integer({ min: -4, max: 4 }),
          warning: fc.boolean(),
        }),
        { maxLength: 8 },
      ),
      fc.array(
        fc.record({
          row: fc.integer({ min: 0, max: 11 }),
          column: fc.integer({ min: 0, max: 11 }),
        }),
        { maxLength: 40 },
      ),
      async (definitions, positions) => {
        const rules = definitions.map((d, index) => ({
          id: String(index),
          range: { startRow: d.r, endRow: d.r + d.h, startColumn: d.c, endColumn: d.c + d.w },
          rule: {
            kind: 'number' as const,
            operator: 'gte' as const,
            first: d.threshold,
            allowBlank: false,
            policy: d.warning ? ('warn' as const) : ('reject' as const),
          },
        }));
        const details: Array<{
          row: number;
          column: number;
          ruleId: string;
          policy: 'warn' | 'reject';
        }> = [];
        let rejected = 0,
          warnings = 0,
          reads = 0;
        for (const cell of positions)
          for (const [index, definition] of definitions.entries()) {
            if (
              cell.row < definition.r ||
              cell.row >= definition.r + definition.h ||
              cell.column < definition.c ||
              cell.column >= definition.c + definition.w ||
              cell.row - cell.column >= definition.threshold
            )
              continue;
            if (definition.warning) warnings++;
            else rejected++;
            if (details.length < 32)
              details.push({
                ...cell,
                ruleId: String(index),
                policy: definition.warning ? 'warn' : 'reject',
              });
          }
        const report = await checkStagedValidation(rules, positions, async (row, column) => {
          reads++;
          return row - column;
        });
        expect(report).toEqual({
          rejected,
          warnings,
          details,
          truncated: rejected + warnings > details.length,
        });
        expect(reads).toBe(
          positions.filter((cell) =>
            definitions.some(
              (d) =>
                cell.row >= d.r &&
                cell.row < d.r + d.h &&
                cell.column >= d.c &&
                cell.column < d.c + d.w,
            ),
          ).length,
        );
        expect(
          checkValidationBatch(
            rules,
            positions.map((cell) => ({ ...cell, value: cell.row - cell.column })),
          ),
        ).toEqual(report);
      },
    ),
    { seed: 712613, numRuns: 1000 },
  );
});

it('can abort during position validation before reaching a late invalid position', async () => {
  const clock = vi.spyOn(performance, 'now');
  let now = 0;
  clock.mockImplementation(() => (now += 10));
  const controller = new AbortController();
  const reason = new Error('stop position validation');
  const positions = Array.from({ length: 65536 }, (_, row) => ({ row, column: 0 }));
  positions[positions.length - 1].column = -1;
  const read = vi.fn(async () => 1);
  try {
    setTimeout(() => controller.abort(reason), 0);
    await expect(checkStagedValidation([], positions, read, controller.signal)).rejects.toBe(
      reason,
    );
    expect(read).not.toHaveBeenCalled();
    await expect(checkStagedValidation([], positions, read)).rejects.toMatchObject({
      code: 'INVALID_ARGUMENT',
      message: 'Invalid validation cell position',
    });
    expect(read).not.toHaveBeenCalled();
  } finally {
    clock.mockRestore();
  }
});

it('yields and can abort while scanning uncovered rule and cell pairs', async () => {
  const clock = vi.spyOn(performance, 'now');
  let now = 0;
  clock.mockImplementation(() => (now += 10));
  const controller = new AbortController();
  const reason = new Error('stop validation');
  const read = vi.fn(async () => 1);
  const rules = Array.from({ length: 100 }, (_, i) => ({
    id: `rule-${i}`,
    range: { startRow: 1000, endRow: 1001, startColumn: 0, endColumn: 1 },
    rule: reject,
  }));
  const positions = Array.from({ length: 1000 }, (_, row) => ({ row, column: 0 }));
  try {
    setTimeout(() => controller.abort(reason), 0);
    await expect(checkStagedValidation(rules, positions, read, controller.signal)).rejects.toBe(
      reason,
    );
    expect(read).not.toHaveBeenCalled();
    const report = await checkStagedValidation(rules, positions, read);
    expect(report).toEqual({ rejected: 0, warnings: 0, details: [], truncated: false });
    expect(read).not.toHaveBeenCalled();
    expect(
      checkValidationBatch(
        rules,
        positions.map((position) => ({ ...position, value: 1 })),
      ),
    ).toEqual(report);
  } finally {
    clock.mockRestore();
  }
});

it('rejects an invalid formula before reading staged cells or invoking the evaluator', async () => {
  const read = vi.fn(async () => 1);
  const evaluate = vi.fn(async (_request: { row: number }) => true);
  const rule = validationRuleSchema.parse({
    kind: 'formula',
    expression: 'A1+',
    allowBlank: false,
    policy: 'reject',
  });
  await expect(
    checkStagedValidation(
      [{ id: 'invalid', range, rule }],
      [{ row: 0, column: 0 }],
      read,
      undefined,
      evaluate,
    ),
  ).rejects.toMatchObject({
    code: 'INVALID_ARGUMENT',
    message: 'Invalid validation formula syntax',
  });
  expect(read).not.toHaveBeenCalled();
  expect(evaluate).not.toHaveBeenCalled();
});

it('applies the blank policy to formula rules without invoking the formula for blank values', async () => {
  const emptyRule = (id: string, allowBlank: boolean, policy: 'reject' | 'warn') => ({
    id,
    range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
    rule: validationRuleSchema.parse({ kind: 'formula', expression: 'TRUE', allowBlank, policy }),
  });
  const read = vi.fn(async (row: number) => [null, '', 5][row] ?? null);
  const evaluate = vi.fn(async (_request: { row: number }) => true);
  const result = await checkStagedValidation(
    [
      emptyRule('required', false, 'reject'),
      emptyRule('warning', false, 'warn'),
      emptyRule('optional', true, 'reject'),
    ],
    [
      { row: 0, column: 0 },
      { row: 1, column: 0 },
      { row: 2, column: 0 },
    ],
    read,
    undefined,
    evaluate,
  );
  expect(result.rejected).toBe(2);
  expect(result.warnings).toBe(2);
  expect(result.details.map(({ row, ruleId }) => [row, ruleId])).toEqual([
    [0, 'required'],
    [0, 'warning'],
    [1, 'required'],
    [1, 'warning'],
  ]);
  expect(evaluate).toHaveBeenCalledTimes(3);
  expect(evaluate.mock.calls.every(([request]) => request.row === 2)).toBe(true);
});

it('evaluates custom rules against real staged cells without replacing the target', async () => {
  const calc = new AsyncCalculation(
    async (sheet, row, column) => {
      if (column !== 0 || row > 1) return;
      return {
        rowId: row,
        columnId: column,
        input: { type: 'number', value: sheet === 'limits' ? 10 : row ? 11 : 5 },
      };
    },
    (name) => (name === 'Limits' ? 'limits' : undefined),
  );
  const rule = validationRuleSchema.parse({
    kind: 'formula',
    expression: 'AND(A1<Limits!$A$1,$A$1=5)',
    allowBlank: false,
    policy: 'reject',
  });
  const reader = calc.validationSession().forSheet('data');
  const result = await checkStagedValidation(
    [{ id: 'limits', range, rule }],
    [
      { row: 0, column: 0 },
      { row: 1, column: 0 },
    ],
    reader.read,
    undefined,
    reader.evaluateFormula,
  );
  expect(result.details).toEqual([{ row: 1, column: 0, ruleId: 'limits', policy: 'reject' }]);
  expect(await calc.value('data', 0, 0)).toBe(5);
  expect(await calc.value('data', 1, 0)).toBe(11);
});

it('isolates staged formula values and parsed expressions from the prior revision cache', async () => {
  let formula = 'B1+1';
  let source = 1;
  const calc = new AsyncCalculation(
    async (_sheet, row, column) => {
      if (row !== 0) return;
      if (column === 0)
        return { rowId: row, columnId: column, input: { type: 'formula', expression: formula } };
      if (column === 1)
        return { rowId: row, columnId: column, input: { type: 'number', value: source } };
    },
    () => undefined,
    undefined,
    async () => ({ kind: 'literal', value: 2 }),
  );
  expect(await calc.value('data', 0, 0)).toBe(2);
  formula = 'B1+5';
  source = 10;
  const staged = calc.validationSession().forSheet('data');
  expect(await staged.read(0, 0)).toBe(15);
  expect(await staged.evaluateFormula({ expression: 'A1=15', rowOffset: 0, columnOffset: 0 })).toBe(
    true,
  );
  expect(await calc.value('data', 0, 0)).toBe(2);
});

it('rejects offset references outside global bounds without reading and respects a zero cache budget', async () => {
  let reads = 0;
  const calc = new AsyncCalculation(
    async (_sheet, row, column) => {
      reads++;
      return { rowId: row, columnId: column, input: { type: 'formula', expression: '1+1' } };
    },
    () => undefined,
    0,
  );
  const staged = calc.validationSession().forSheet('data');
  expect(
    await staged.evaluateFormula({ expression: 'A1', rowOffset: -1, columnOffset: 0 }),
  ).toEqual({ error: '#REF!' });
  expect(
    await staged.evaluateFormula({ expression: 'A1', rowOffset: 100_000, columnOffset: 0 }),
  ).toEqual({ error: '#REF!' });
  expect(
    await staged.evaluateFormula({ expression: 'A1', rowOffset: 0, columnOffset: 1000 }),
  ).toEqual({ error: '#REF!' });
  expect(reads).toBe(0);
  expect(await staged.read(0, 0)).toBe(2);
  expect(await staged.read(0, 0)).toBe(2);
  expect(reads).toBe(2);
});

it('shares the staged calculation budget across candidate reads, rules and sheets', async () => {
  const reads: string[] = [];
  const calc = new AsyncCalculation(
    async (sheet) => {
      reads.push(sheet);
      return { rowId: 0, columnId: 0, input: { type: 'number', value: 12 } };
    },
    () => undefined,
  );
  const session = calc.validationSession(undefined, 3);
  expect(await session.forSheet('first').read(0, 0)).toBe(12);
  expect(
    await session
      .forSheet('first')
      .evaluateFormula({ expression: 'TRUE', rowOffset: 0, columnOffset: 0 }),
  ).toBe(true);
  await expect(
    checkStagedValidation(
      [{ id: 'second', range, rule: reject }],
      [{ row: 0, column: 0 }],
      session.forSheet('second').read,
    ),
  ).rejects.toThrow('Validation calculation exceeded');
  expect(reads).toEqual(['first']);
});

it('checks custom formulas relative to the rule origin and accepts only boolean true', async () => {
  const rule = validationRuleSchema.parse({
    kind: 'formula',
    expression: 'A1>0',
    allowBlank: true,
    policy: 'warn',
  });
  const rules = [{ id: 'formula', range: { ...range, endRow: 4 }, rule }];
  const cells = [0, 1, 2, 3].map((row) => ({ row, column: 0 }));
  const evaluate = async ({
    row,
    rowOffset,
    columnOffset,
    expression,
  }: {
    row: number;
    rowOffset: number;
    columnOffset: number;
    expression: string;
  }) => {
    expect([rowOffset, columnOffset, expression]).toEqual([row, 0, 'A1>0']);
    expect(row).not.toBe(3);
    return row === 0 ? true : row === 1 ? 1 : { error: '#REF!' };
  };
  const read = async (row: number) => (row === 3 ? null : 1);
  const report = await checkStagedValidation(rules, cells, read, undefined, evaluate);
  expect(report.details.map(({ row }) => row)).toEqual([1, 2]);
  expect(report.warnings).toBe(2);
  await expect(checkStagedValidation(rules, cells, read)).rejects.toThrow(
    'formula reader required',
  );
  await expect(
    checkStagedValidation(rules, cells, read, undefined, async () => ({ error: '#LIMIT!' })),
  ).rejects.toThrow('Validation calculation exceeded');
});

it.each(['reject', 'warn'] as const)(
  'does not accept an error-valued cell under a true formula rule with %s policy',
  async (policy) => {
    const rule = validationRuleSchema.parse({
      kind: 'formula',
      expression: 'TRUE',
      allowBlank: true,
      policy,
    });
    let evaluations = 0;
    const report = await checkStagedValidation(
      [{ id: 'f', range, rule }],
      [{ row: 0, column: 0 }],
      async () => ({ error: '#DIV/0!' }),
      undefined,
      async () => {
        evaluations++;
        return true;
      },
    );
    expect(report.rejected).toBe(policy === 'reject' ? 1 : 0);
    expect(report.warnings).toBe(policy === 'warn' ? 1 : 0);
    expect(report.details).toEqual([{ row: 0, column: 0, ruleId: 'f', policy }]);
    expect(evaluations).toBe(0);
    await expect(
      checkStagedValidation(
        [{ id: 'f', range, rule }],
        [{ row: 0, column: 0 }],
        async () => ({ error: '#LIMIT!' }),
        undefined,
        async () => true,
      ),
    ).rejects.toThrow('Validation calculation exceeded');
  },
);

it('rejects a staged formula cell whose calculation produces an error', async () => {
  const calc = new AsyncCalculation(
    async (_sheet, row, column) => ({
      rowId: row,
      columnId: column,
      input: { type: 'formula', expression: '1/0' },
    }),
    () => undefined,
  );
  const session = calc.validationSession().forSheet('data');
  const rule = validationRuleSchema.parse({
    kind: 'formula',
    expression: 'TRUE',
    allowBlank: false,
    policy: 'reject',
  });
  const report = await checkStagedValidation(
    [{ id: 'f', range, rule }],
    [{ row: 0, column: 0 }],
    session.read,
    undefined,
    session.evaluateFormula,
  );
  expect(report.rejected).toBe(1);
  expect(report.warnings).toBe(0);
});

it('reads only covered staged cells once even when rules overlap', async () => {
  const rules = [
    { id: 'required', range, rule: reject },
    { id: 'advice', range, rule: warn },
  ];
  const cells = [
    { row: 0, column: 0, value: 7 },
    { row: 3, column: 0, value: 1 },
  ];
  const reads: number[] = [];
  const result = await checkStagedValidation(rules, cells, async (row) => {
    reads.push(row);
    return 7;
  });
  expect(reads).toEqual([0]);
  expect(result).toEqual(checkValidationBatch(rules, cells));
});

it('keeps candidate coordinates and rules stable when the caller mutates them during an async read', async () => {
  const rules = [{ id: 'required', range: { ...range }, rule: { ...reject } }];
  const positions = [
    { row: 0, column: 0 },
    { row: 1, column: 0 },
  ];
  const reads: number[] = [];
  let release!: () => void;
  let started!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const entered = new Promise<void>((resolve) => {
    started = resolve;
  });
  const checking = checkStagedValidation(rules, positions, async (row) => {
    reads.push(row);
    if (row === 0) {
      started();
      await blocked;
    }
    return row === 0 ? 0 : 20;
  });
  await entered;
  positions[0].row = 3;
  positions[1].row = 3;
  positions.push({ row: 3, column: 0 });
  rules[0].range.endRow = 1;
  rules[0].rule.policy = 'warn';
  release();
  const report = await checking;
  expect(reads).toEqual([0, 1]);
  expect(report).toEqual({
    rejected: 1,
    warnings: 0,
    details: [{ row: 0, column: 0, ruleId: 'required', policy: 'reject' }],
    truncated: false,
  });
});

it('returns no partial report on cancellation or a late calculation limit', async () => {
  const rules = [{ id: 'required', range, rule: reject }];
  const cells = [
    { row: 0, column: 0 },
    { row: 1, column: 0 },
  ];
  const controller = new AbortController();
  const reads: number[] = [];
  await expect(
    checkStagedValidation(
      rules,
      cells,
      async (row) => {
        reads.push(row);
        controller.abort();
        return 0;
      },
      controller.signal,
    ),
  ).rejects.toHaveProperty('name', 'AbortError');
  expect(reads).toEqual([0]);
  await expect(
    checkStagedValidation(rules, cells, async (row) => (row === 0 ? 0 : { error: '#LIMIT!' })),
  ).rejects.toThrow('Validation calculation exceeded');
});

it.each(['read', 'formula'] as const)(
  'stops waiting when cancellation interrupts a pending %s operation',
  async (pending) => {
    const controller = new AbortController();
    let started!: () => void;
    let rejectPending!: (error: Error) => void;
    const entered = new Promise<void>((resolve) => {
      started = resolve;
    });
    const waiting = new Promise<never>((_, reject) => {
      rejectPending = reject;
    });
    const rule =
      pending === 'read'
        ? reject
        : validationRuleSchema.parse({
            kind: 'formula',
            expression: 'TRUE',
            allowBlank: false,
            policy: 'reject',
          });
    const report = checkStagedValidation(
      [{ id: 'check', range, rule }],
      [{ row: 0, column: 0 }],
      async () => {
        if (pending === 'read') {
          started();
          return waiting;
        }
        return 1;
      },
      controller.signal,
      async () => {
        started();
        return waiting;
      },
    );
    await entered;
    controller.abort();
    await expect(report).rejects.toHaveProperty('name', 'AbortError');
    rejectPending(new Error('late failure'));
    await Promise.resolve();
  },
);

it('reports overlapping reject and warning rules with bounded details and no mutation', () => {
  const rules = [
    { id: 'required', range, rule: reject },
    { id: 'advice', range, rule: warn },
  ];
  const cells = [
    { row: 0, column: 0, value: 7 },
    { row: 1, column: 0, value: null },
    { row: 2, column: 0, value: 3 },
  ];
  const before = structuredClone({ rules, cells });
  expect(checkValidationBatch(rules, cells)).toEqual({
    rejected: 2,
    warnings: 2,
    truncated: false,
    details: [
      { row: 0, column: 0, ruleId: 'required', policy: 'reject' },
      { row: 0, column: 0, ruleId: 'advice', policy: 'warn' },
      { row: 1, column: 0, ruleId: 'required', policy: 'reject' },
      { row: 1, column: 0, ruleId: 'advice', policy: 'warn' },
    ],
  });
  expect({ rules, cells }).toEqual(before);
});

it('counts every failure even when details are truncated', () => {
  const entries = Array.from({ length: 40 }, (_, index) => ({
    id: `r${index}`,
    range,
    rule: reject,
  }));
  const result = checkValidationBatch(entries, [{ row: 0, column: 0, value: 0 }]);
  expect(result.rejected).toBe(40);
  expect(result.warnings).toBe(0);
  expect(result.details).toHaveLength(32);
  expect(result.truncated).toBe(true);
});

it('rejects invalid geometry, duplicate identities and oversized work without returning partial results', () => {
  expect(() =>
    checkValidationBatch([{ id: 'x', range: { ...range, endRow: 0 }, rule: reject }], []),
  ).toThrow('Invalid validation rule range');
  expect(() =>
    checkValidationBatch(
      [
        { id: 'x', range, rule: reject },
        { id: 'x', range, rule: warn },
      ],
      [],
    ),
  ).toThrow('Invalid validation rule range');
  expect(() => checkValidationBatch([], [{ row: -1, column: 0, value: 1 }])).toThrow(
    'Invalid validation cell',
  );
  expect(() =>
    checkValidationBatch(
      [{ id: 'x', range, rule: { ...reject, first: Infinity } as typeof reject }],
      [],
    ),
  ).toThrow('Invalid validation rule');
  expect(() =>
    checkValidationBatch(
      [{ id: 'x', range, rule: reject }],
      [{ row: 0, column: 0, value: { error: '#LIMIT!' } }],
    ),
  ).toThrow('Validation calculation exceeded');
  expect(
    checkValidationBatch(
      [{ id: 'x', range, rule: reject }],
      [{ row: 9, column: 0, value: { error: '#LIMIT!' } }],
    ).rejected,
  ).toBe(0);
  const longList = validationRuleSchema.parse({
    kind: 'list',
    values: ['x'.repeat(32768)],
    allowBlank: false,
    policy: 'reject',
  });
  expect(() =>
    checkValidationBatch(
      [{ id: 'x', range, rule: longList }],
      Array.from({ length: 31 }, () => ({ row: 0, column: 0, value: 'x'.repeat(32768) })),
    ),
  ).toThrow('Validation work budget exceeded');
  const metadataHeavy = Array.from({ length: 31 }, (_, index) => ({
    id: `long${index}`,
    range,
    rule: longList,
  }));
  expect(() => checkValidationBatch(metadataHeavy, [])).toThrow('Validation work budget exceeded');
  const entries = Array.from({ length: 1001 }, (_, index) => ({
    id: `r${index}`,
    range,
    rule: reject,
  }));
  const cells = Array.from({ length: 1000 }, () => ({ row: 0, column: 0, value: 0 }));
  expect(() => checkValidationBatch(entries, cells)).toThrow('Validation work budget exceeded');
});

it('checks a large list against a whole column without rescanning its entries', () => {
  const values = Array.from({ length: 1000 }, (_, i) => `choice-${i}`);
  const rule = validationRuleSchema.parse({
    kind: 'list',
    values,
    allowBlank: false,
    policy: 'reject',
  });
  const cells = Array.from({ length: 65536 }, (_, row) => ({
    row,
    column: 0,
    value: values[row % values.length],
  }));
  cells[0].value = 'Choice-0';
  cells[65535].value = '';
  const result = checkValidationBatch(
    [{ id: 'choices', range: { ...range, endRow: cells.length }, rule }],
    cells,
  );
  expect(result.details.map(({ row }) => row)).toEqual([0, 65535]);
  expect(result.rejected).toBe(2);
});

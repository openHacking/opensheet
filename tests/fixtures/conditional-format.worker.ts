import { AsyncCalculation } from '../../packages/core/src/async-calculation.js';
import {
  conditionalViewportWithFormulas,
  validateConditionalRules,
} from '../../packages/core/src/conditional-format.js';

let controller: AbortController | undefined;
self.onmessage = async (event: MessageEvent<'run' | 'wait' | 'cancel'>) => {
  if (event.data === 'cancel') {
    controller?.abort();
    return;
  }
  controller = new AbortController();
  try {
    const calc = new AsyncCalculation(
      async (sheet, row, column) =>
        column === 0 && row < 2
          ? {
              rowId: row,
              columnId: column,
              input: { type: 'number', value: sheet === 'limits' ? 10 : row ? 11 : 5 },
            }
          : undefined,
      (name) => (name === 'Limits' ? 'limits' : undefined),
    );
    const rules = validateConditionalRules(
      [
        {
          id: 'formula',
          range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
          condition: { kind: 'formula', expression: 'AND(A1>Limits!$A$1,$A$1=5)' },
          style: { bold: true },
          stopIfTrue: false,
        },
      ],
      10,
      10,
    );
    const cells = [
      { row: 0, column: 0, value: 5, base: {} },
      { row: 1, column: 0, value: 11, base: {} },
    ];
    const reader = calc.validationSession(controller.signal).forSheet('data');
    const evaluate =
      event.data === 'wait'
        ? () => {
            self.postMessage({ kind: 'waiting' });
            return new Promise<never>(() => {});
          }
        : reader.evaluateFormula;
    const styles = await conditionalViewportWithFormulas(rules, cells, evaluate, controller.signal);
    self.postMessage({ kind: 'result', styles });
  } catch (error) {
    self.postMessage({ kind: 'error', name: error instanceof Error ? error.name : String(error) });
  }
};

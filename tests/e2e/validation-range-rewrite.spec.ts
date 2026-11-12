import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('staged validation and conditional rules preserve partially deleted formula ranges in a Worker', async ({
  page,
}) => {
  const path = (name: string) =>
    `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
  await page.route('**/range-rule-rewrite.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { transformValidationEntries } from ${JSON.stringify(path('validation-model'))};
      import { transformConditionalRules } from ${JSON.stringify(path('conditional-format'))};
      try {
        const range = {startRow:6,endRow:8,startColumn:0,endColumn:1};
        const expression = 'AND(SUM(A1:A5)>0,SUM($A$5:$A$2)>0,SUM(A1:A2)>0)';
        const input = [{id:'v',range,rule:{kind:'formula',expression,allowBlank:false,policy:'reject'}}];
        const conditional = [{id:'c',range,condition:{kind:'formula',expression},style:{bold:true},stopIfTrue:false}];
        const v = transformValidationEntries(input,10,4,'row','delete',0,2,'Data');
        const c = transformConditionalRules(conditional,10,4,'row','delete',0,2,'Data');
        self.postMessage({validation:v[0],conditional:c[0],original:input[0].rule.expression});
      } catch (error) { self.postMessage({error:String(error)}); }
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/range-rule-rewrite.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker timed out'));
        }, 10000);
        worker.onerror = (event) => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error(event.message));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timer);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result.error).toBeUndefined();
  expect(result.validation.range).toEqual({ startRow: 4, endRow: 6, startColumn: 0, endColumn: 1 });
  expect(result.conditional.range).toEqual(result.validation.range);
  expect(result.validation.rule.expression).toBe(result.conditional.condition.expression);
  expect(result.validation.rule.expression).toContain('A1:A3');
  expect(result.validation.rule.expression).toContain('$A$3:$A$1');
  expect(result.validation.rule.expression).toContain('#REF!');
  expect(result.original).toContain('A1:A5');
});

import { writeFileSync } from 'node:fs';
import { functionDefinitions } from '../packages/formula/dist/index.js';

const rows = functionDefinitions.map(
  ({ name, signature, minArgs, maxArgs, evaluation, description }) =>
    `| ${name} | \`${signature}\` | ${minArgs}–${maxArgs} | ${evaluation} | ${description} |`,
);
writeFileSync(
  'docs/formula-functions.md',
  [
    '# Formula function reference',
    '',
    'Generated from the runtime function registry by `pnpm formula:docs` after building packages. Do not edit this table manually. See [semantics and limitations](formulas.md) before assuming Excel compatibility.',
    '',
    '| Function | Signature | Arguments | Evaluation | Description |',
    '| --- | --- | --- | --- | --- |',
    ...rows,
    '',
  ].join('\n'),
);

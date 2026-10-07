import { toInput, type CellInput } from './types.js';

export function parseInput(text: string): CellInput {
  if (text.startsWith('=')) return { type: 'formula', expression: text.slice(1) };
  if (text.startsWith("'")) return { type: 'string', value: text.slice(1) };
  if (text === '') return { type: 'blank' };
  if (/^(true|false)$/i.test(text))
    return { type: 'boolean', value: text.toLowerCase() === 'true' };
  if (
    /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(text) &&
    text.replace(/[^\d]/g, '').length <= 15
  )
    return toInput(Number(text));
  return { type: 'string', value: text };
}

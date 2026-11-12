const explanations: Record<string, string> = {
  '#DIV/0!':
    'A calculation divided by zero or an average had no numeric values. Check divisors and referenced cells.',
  '#VALUE!':
    'A value, argument, or formula expression could not be used in this calculation. Check argument types and referenced cells.',
  '#REF!':
    'A reference is invalid or points outside the available sheet. Check references and deleted rows or columns.',
  '#NAME?': 'A function or name is not recognized. Check spelling and the supported function list.',
  '#N/A':
    'A requested value was not available, for example when a lookup found no match. Check the lookup inputs and referenced cells.',
  '#NUM!':
    'A numeric result or argument is outside the supported range or domain. Check numeric inputs and referenced cells.',
  '#CYCLE!':
    'This formula depends on itself through one or more references. Check the dependency chain.',
  '#LIMIT!':
    'Calculation exceeded a resource limit. Simplify the formula or reduce the referenced workload.',
};
/** General guidance; an error can originate in a referenced cell rather than this formula. */
export function formulaErrorExplanation(code: string): string | undefined {
  return Object.hasOwn(explanations, code) ? explanations[code] : undefined;
}

export class FormulaError extends Error {
  constructor(public code: string) {
    super(code);
  }
}
export const fail = (code = '#VALUE!'): never => {
  throw new FormulaError(code);
};

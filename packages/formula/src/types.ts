export type FormulaValue = string | number | boolean | null | { error: string };
export type Reference = {
  kind: 'ref';
  row: number;
  column: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
  sheet?: string;
};
export type AST =
  | Reference
  | { kind: 'literal'; value: FormulaValue }
  | { kind: 'range'; start: Reference; end: Reference }
  | { kind: 'unary'; op: string; value: AST }
  | { kind: 'binary'; op: string; left: AST; right: AST }
  | { kind: 'call'; name: string; args: AST[] };

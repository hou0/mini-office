export const ERROR_CODES = {
  DIV_ZERO: '#DIV/0!',
  VALUE: '#VALUE!',
  REF: '#REF!',
  NAME: '#NAME?',
  NUM: '#NUM!',
  NA: '#N/A',
  NULL: '#NULL!',
  CIRC: '#CIRC!',
} as const;

export type ErrorCode = (typeof ERROR_CODES)[keyof typeof ERROR_CODES];

export class FormulaError extends Error {
  readonly code: ErrorCode;

  constructor(code: ErrorCode, message?: string) {
    super(message ?? code);
    this.name = 'FormulaError';
    this.code = code;
  }
}

export type Scalar = number | string | boolean | null;

export interface RangeValue {
  kind: 'range';
  sheetName: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  rows: number;
  cols: number;
  cells: EvalValue[][];
}

export type EvalValue = Scalar | FormulaError | RangeValue;

export function isRange(value: EvalValue): value is RangeValue {
  return !!value && typeof value === 'object' && (value as RangeValue).kind === 'range';
}

export function isError(value: EvalValue): value is FormulaError {
  return value instanceof FormulaError;
}

export function error(code: ErrorCode): FormulaError {
  return new FormulaError(code);
}

export function isErrorString(value: unknown): value is string {
  return typeof value === 'string' && /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|N\/A|CIRC!)$/i.test(value);
}

export function toErrorString(value: unknown): string | undefined {
  return value instanceof FormulaError ? value.code : isErrorString(value) ? value.toUpperCase() : undefined;
}

/** Scalarize a range value (implicit top-left intersection). */
export function unwrap(value: EvalValue): EvalValue {
  if (!isRange(value)) return value;
  return value.cells[0]?.[0] ?? null;
}

export function flattenRange(value: EvalValue): EvalValue[] {
  if (!isRange(value)) return [value];
  const result: EvalValue[] = [];
  for (const row of value.cells) for (const cell of row) result.push(cell);
  return result;
}

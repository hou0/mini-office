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

/** 功能：复用同一个比较器，避免每次文本比较都重新创建 Intl.Collator；时间：2026-10-27；作者：lq。 */
const textCollator = new Intl.Collator('zh-CN', { sensitivity: 'accent' });

/** 与 String.prototype.localeCompare(value, 'zh-CN', { sensitivity: 'accent' }) 语义一致，但复用比较器。 */
export function compareText(left: string, right: string): number {
  if (left === right) return 0;
  return textCollator.compare(left, right);
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
  /** 引用自整列（如 $A:$A），标量语境下按公式所在行做隐式交集。 */
  wholeColumn?: boolean;
  /** 引用自整行（如 $1:$1），标量语境下按公式所在列做隐式交集。 */
  wholeRow?: boolean;
  /** 单元格矩阵；公式引擎惰性生成，只有旧接口直接读取时才会整体求值。 */
  cells: EvalValue[][];
  /** 惰性读取单个单元格，参数为工作表绝对行列号。 */
  cellAt?: (row: number, col: number) => EvalValue;
}

/** 区域内的数值统计，供 SUM/MAX 等聚合函数复用，避免重复展开整个区域。 */
export interface RangeStats {
  /** 区域内数值单元格个数。 */
  count: number;
  sum: number;
  max: number;
  min: number;
  product: number;
  /** 区域内非空单元格个数（含文本与布尔值）。 */
  nonEmpty: number;
  error?: FormulaError;
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
  if (value.cellAt) return value.cellAt(value.r1, value.c1);
  return value.cells[0]?.[0] ?? null;
}

/** 功能：按 Excel 隐式交集规则把区域转为标量：整列取公式所在行、整行取公式所在列，其余取左上角；时间：2026-10-28；作者：lq。 */
export function intersectScalar(value: EvalValue, currentRow: number, currentCol: number): EvalValue {
  if (!isRange(value)) return value;
  const row = value.wholeColumn ? currentRow : value.r1;
  const col = value.wholeRow ? currentCol : value.c1;
  if (row < value.r1 || row > value.r2 || col < value.c1 || col > value.c2) return error(ERROR_CODES.REF);
  if (value.cellAt) return value.cellAt(row, col);
  return value.cells[row - value.r1]?.[col - value.c1] ?? null;
}

/** 惰性展开区域内的单元格，避免为了一次取值而整体物化矩形。 */
export function flattenRange(value: EvalValue): EvalValue[] {
  if (!isRange(value)) return [value];
  const result: EvalValue[] = [];
  if (value.cellAt) {
    for (let row = value.r1; row <= value.r2; row += 1) {
      for (let col = value.c1; col <= value.c2; col += 1) result.push(value.cellAt(row, col));
    }
    return result;
  }
  for (const row of value.cells) for (const cell of row) result.push(cell);
  return result;
}

/** 按区域内的相对行列读取单元格，不物化整个区域。 */
export function rangeCell(range: RangeValue, row: number, col: number): EvalValue {
  if (range.cellAt) return range.cellAt(range.r1 + row, range.c1 + col);
  return range.cells[row]?.[col] ?? null;
}

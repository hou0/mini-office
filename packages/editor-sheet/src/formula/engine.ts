import type {
  Cell,
  CellScalar,
  ComputedCell,
  SheetModel,
  WorkbookModel,
} from '@mini-office/document-model';
import {
  ERROR_CODES,
  FormulaError,
  error,
  flattenRange,
  isError,
  isRange,
  unwrap,
  type EvalValue,
  type RangeValue,
} from './types';
import { parseFormula, type Expr } from './parser';
import { getFunction, type FunctionContext } from './functions';
import { formatCellValue } from '../format';

const MAX_RANGE_CELLS = 300_000;

function isErrorString(value: string): boolean {
  return /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|N\/A|CIRC!)$/i.test(value);
}

function scalarOf(value: EvalValue): CellScalar {
  const scalar = unwrap(value);
  if (isRange(scalar)) return scalarOf(scalar);
  if (isError(scalar)) return scalar.code;
  return scalar;
}

export function evaluateFormula(
  formula: string,
  workbook: WorkbookModel,
  sheetName: string,
  row = 0,
  col = 0,
): EvalValue {
  const sheet = workbook.sheets.find((item) => item.name.toLowerCase() === sheetName.toLowerCase());
  if (!sheet) return error(ERROR_CODES.REF);
  const evaluator = new WorkbookEvaluator(workbook);
  return evaluator.evaluateFormulaAt(formula, sheet, row, col);
}

class WorkbookEvaluator {
  private readonly workbook: WorkbookModel;
  private readonly sheetById = new Map<string, SheetModel>();
  private readonly sheetByName = new Map<string, SheetModel>();
  private readonly cache = new Map<string, EvalValue>();
  private readonly stack = new Set<string>();
  private readonly nameStack = new Set<string>();
  private readonly now = new Date();
  private readonly date1904: boolean;

  constructor(workbook: WorkbookModel) {
    this.workbook = workbook;
    this.date1904 = workbook.date1904 === true;
    for (const sheet of workbook.sheets) {
      this.sheetById.set(sheet.id, sheet);
      this.sheetByName.set(sheet.name.toLowerCase(), sheet);
    }
  }

  private resolveSheet(name?: string): SheetModel | undefined {
    if (!name) return undefined;
    return this.sheetByName.get(name.toLowerCase());
  }

  private rawCellValue(sheet: SheetModel, row: number, col: number): EvalValue {
    if (row < 0 || col < 0 || row > 1_048_575 || col > 16_383) {
      return error(ERROR_CODES.REF);
    }
    const key = `${row}:${col}`;
    return this.evaluateCell(sheet.id, key);
  }

  evaluateCell(sheetId: string, key: string): EvalValue {
    const cacheKey = `${sheetId}!${key}`;
    const cached = this.cache.get(cacheKey);
    if (cached !== undefined) return cached;
    if (this.stack.has(cacheKey)) return error(ERROR_CODES.CIRC);
    const sheet = this.sheetById.get(sheetId);
    if (!sheet) return error(ERROR_CODES.REF);
    const cell = sheet.cells[key];
    if (!cell) {
      this.cache.set(cacheKey, null);
      return null;
    }
    this.stack.add(cacheKey);
    let value: EvalValue;
    try {
      if (cell.formula) {
        const [row, col] = key.split(':').map(Number);
        const ast = parseFormula(cell.formula);
        value = this.evaluateExpr(ast, sheet, row, col);
        if (isRange(value)) value = unwrap(value);
      } else {
        value = cell.value;
        if (typeof value === 'string' && isErrorString(value)) {
          value = new FormulaError(value.toUpperCase() as FormulaError['code']);
        }
      }
    } catch (caught) {
      value =
        caught instanceof FormulaError
          ? caught
          : new FormulaError(ERROR_CODES.VALUE, caught instanceof Error ? caught.message : String(caught));
    } finally {
      this.stack.delete(cacheKey);
    }
    this.cache.set(cacheKey, value);
    return value;
  }

  evaluateFormulaAt(formula: string, sheet: SheetModel, row: number, col: number): EvalValue {
    try {
      return this.evaluateExpr(parseFormula(formula), sheet, row, col);
    } catch (caught) {
      return caught instanceof FormulaError ? caught : error(ERROR_CODES.VALUE);
    }
  }

  private createRange(
    sheetName: string | undefined,
    r1: number,
    c1: number,
    r2: number,
    c2: number,
    currentSheet: SheetModel,
    wholeColumn: boolean,
    wholeRow: boolean,
  ): EvalValue {
    const sheet = sheetName ? this.resolveSheet(sheetName) : currentSheet;
    if (!sheet) return error(ERROR_CODES.REF);
    let top = Math.min(r1, r2);
    let bottom = Math.max(r1, r2);
    let left = Math.min(c1, c2);
    let right = Math.max(c1, c2);
    if (wholeColumn) {
      top = 0;
      bottom = Math.max(0, Math.min(bottom, sheet.dimension.rows - 1));
    }
    if (wholeRow) {
      left = 0;
      right = Math.max(0, Math.min(right, sheet.dimension.cols - 1));
    }
    top = Math.max(0, top);
    left = Math.max(0, left);
    bottom = Math.min(1_048_575, bottom);
    right = Math.min(16_383, right);
    const rows = bottom - top + 1;
    const cols = right - left + 1;
    if (rows <= 0 || cols <= 0) return error(ERROR_CODES.REF);
    if (rows * cols > MAX_RANGE_CELLS) return error(ERROR_CODES.NUM);
    const cells: EvalValue[][] = [];
    for (let row = top; row <= bottom; row += 1) {
      const line: EvalValue[] = [];
      for (let col = left; col <= right; col += 1) {
        line.push(this.rawCellValue(sheet, row, col));
      }
      cells.push(line);
    }
    return {
      kind: 'range',
      sheetName: sheet.name,
      r1: top,
      c1: left,
      r2: bottom,
      c2: right,
      rows,
      cols,
      cells,
    };
  }

  private evaluateExpr(expr: Expr, sheet: SheetModel, row: number, col: number): EvalValue {
    switch (expr.type) {
      case 'number':
        return expr.value;
      case 'string':
        return expr.value;
      case 'boolean':
        return expr.value;
      case 'error':
        return new FormulaError(expr.value);
      case 'empty':
        return null;
      case 'ref': {
        const target = expr.sheet ? this.resolveSheet(expr.sheet) : sheet;
        if (!target) return error(ERROR_CODES.REF);
        return this.rawCellValue(target, expr.row, expr.col);
      }
      case 'range':
        return this.createRange(
          expr.sheet,
          expr.r1,
          expr.c1,
          expr.r2,
          expr.c2,
          sheet,
          !!expr.wholeColumn,
          !!expr.wholeRow,
        );
      case 'name': {
        const defined = this.workbook.definedNames?.[expr.name] ?? this.workbook.definedNames?.[expr.name.toUpperCase()];
        if (!defined) return error(ERROR_CODES.NAME);
        if (this.nameStack.has(expr.name)) return error(ERROR_CODES.CIRC);
        this.nameStack.add(expr.name);
        try {
          const definition = defined.startsWith('=') ? defined : `=${defined}`;
          return this.evaluateExpr(parseFormula(definition), sheet, row, col);
        } catch {
          return error(ERROR_CODES.NAME);
        } finally {
          this.nameStack.delete(expr.name);
        }
      }
      case 'unary': {
        const value = this.toNumber(this.evaluateExpr(expr.expr, sheet, row, col));
        if (isError(value)) return value;
        return expr.op === '-' ? -value : value;
      }
      case 'percent': {
        const value = this.toNumber(this.evaluateExpr(expr.expr, sheet, row, col));
        if (isError(value)) return value;
        return value / 100;
      }
      case 'binary':
        return this.evaluateBinary(expr, sheet, row, col);
      case 'func': {
        const definition = getFunction(expr.name);
        if (!definition) return error(ERROR_CODES.NAME);
        if (expr.args.length < definition.minArgs || expr.args.length > definition.maxArgs) {
          return error(ERROR_CODES.VALUE);
        }
        const context: FunctionContext = {
          workbook: this.workbook,
          sheet,
          sheetName: sheet.name,
          date1904: this.date1904,
          now: this.now,
          currentRow: row,
          currentCol: col,
          evalArg: (arg) => this.evaluateExpr(arg, sheet, row, col),
          createRange: (sheetName, r1, c1, r2, c2) => {
            const result = this.createRange(sheetName, r1, c1, r2, c2, sheet, false, false);
            if (isError(result)) throw result;
            return result as RangeValue;
          },
          resolveSheet: (name) => (name ? this.resolveSheet(name) : sheet),
        };
        try {
          const result = definition.call(expr.args, context);
          return isRange(result) ? unwrap(result) : result;
        } catch (caught) {
          if (caught instanceof FormulaError) return caught;
          return error(ERROR_CODES.VALUE);
        }
      }
      default:
        return error(ERROR_CODES.NAME);
    }
  }

  private evaluateBinary(
    expr: Extract<Expr, { type: 'binary' }>,
    sheet: SheetModel,
    row: number,
    col: number,
  ): EvalValue {
    const leftRaw = this.evaluateExpr(expr.left, sheet, row, col);
    const rightRaw = this.evaluateExpr(expr.right, sheet, row, col);
    const left = unwrap(leftRaw);
    const right = unwrap(rightRaw);
    if (isError(left)) return left;
    if (isError(right)) return right;
    if (expr.op === '&') {
      return `${this.toText(left)}${this.toText(right)}`;
    }
    if (['=', '<>', '<', '>', '<=', '>='].includes(expr.op)) {
      return this.compare(left, right, expr.op);
    }
    const a = this.toNumber(left);
    const b = this.toNumber(right);
    if (isError(a)) return a;
    if (isError(b)) return b;
    switch (expr.op) {
      case '+':
        return a + b;
      case '-':
        return a - b;
      case '*':
        return a * b;
      case '/':
        return b === 0 ? error(ERROR_CODES.DIV_ZERO) : a / b;
      case '^': {
        const result = a ** b;
        return Number.isFinite(result) ? result : error(ERROR_CODES.NUM);
      }
      default:
        return error(ERROR_CODES.VALUE);
    }
  }

  private compare(left: EvalValue, right: EvalValue, op: string): boolean {
    if (isError(left)) return false;
    if (isError(right)) return false;
    let comparison = 0;
    if (left === null && right === null) comparison = 0;
    else if (left === null) comparison = typeof right === 'string' ? -1 : -1;
    else if (right === null) comparison = 1;
    else if (typeof left === 'number' || typeof right === 'number' || typeof left === 'boolean' || typeof right === 'boolean') {
      const a = this.toNumber(left);
      const b = this.toNumber(right);
      if (isError(a) || isError(b)) {
        comparison = this.toText(left).localeCompare(this.toText(right), 'zh-CN', { sensitivity: 'accent' });
      } else {
        comparison = a - b;
      }
    } else {
      comparison = this.toText(left).localeCompare(this.toText(right), 'zh-CN', { sensitivity: 'accent' });
    }
    switch (op) {
      case '=':
        return comparison === 0;
      case '<>':
        return comparison !== 0;
      case '<':
        return comparison < 0;
      case '>':
        return comparison > 0;
      case '<=':
        return comparison <= 0;
      case '>=':
        return comparison >= 0;
      default:
        return false;
    }
  }

  private toNumber(value: EvalValue): number | FormulaError {
    const scalar = unwrap(value);
    if (isError(scalar)) return scalar;
    if (scalar === null || scalar === undefined) return 0;
    if (typeof scalar === 'number') return Number.isFinite(scalar) ? scalar : error(ERROR_CODES.NUM);
    if (typeof scalar === 'boolean') return scalar ? 1 : 0;
    if (typeof scalar === 'string') {
      const text = scalar.trim();
      if (!text) return 0;
      if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$/.test(text)) {
        return text.endsWith('%') ? Number(text.slice(0, -1)) / 100 : Number(text);
      }
      return error(ERROR_CODES.VALUE);
    }
    return error(ERROR_CODES.VALUE);
  }

  private toText(value: EvalValue): string {
    const scalar = unwrap(value);
    if (isError(scalar)) return scalar.code;
    if (scalar === null || scalar === undefined) return '';
    if (typeof scalar === 'boolean') return scalar ? 'TRUE' : 'FALSE';
    if (typeof scalar === 'number') return formatCellValue(scalar, undefined, this.date1904);
    return String(scalar);
  }

  calculate(): WorkbookModel {
    const computed: Record<string, Record<string, ComputedCell>> = {};
    for (const sheet of this.workbook.sheets) {
      const sheetComputed: Record<string, ComputedCell> = {};
      const keys = Object.keys(sheet.cells);
      for (const key of keys) {
        const value = this.evaluateCell(sheet.id, key);
        const cell = sheet.cells[key];
        const style = cell.styleId ? this.workbook.styles[cell.styleId] : undefined;
        const scalar = scalarOf(value);
        const display = isError(value)
          ? value.code
          : formatCellValue(scalar, style, this.date1904);
        sheetComputed[key] = {
          value: scalar,
          display,
          error: isError(value) ? value.code : undefined,
        };
      }
      computed[sheet.id] = sheetComputed;
    }
    return { ...this.workbook, computed };
  }
}

export function recalculateWorkbook(workbook: WorkbookModel): WorkbookModel {
  return new WorkbookEvaluator(workbook).calculate();
}

export function evaluateCellValue(
  workbook: WorkbookModel,
  sheetId: string,
  row: number,
  col: number,
): EvalValue {
  const evaluator = new WorkbookEvaluator(workbook);
  return evaluator.evaluateCell(sheetId, `${row}:${col}`);
}

export function formulaValueToCell(value: EvalValue): CellScalar {
  return scalarOf(value);
}

export function collectFormulaCells(workbook: WorkbookModel): Array<{ sheetId: string; key: string; formula: string }> {
  const result: Array<{ sheetId: string; key: string; formula: string }> = [];
  for (const sheet of workbook.sheets) {
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (cell.formula) result.push({ sheetId: sheet.id, key, formula: cell.formula });
    }
  }
  return result;
}

export function isFormulaCell(cell: Cell | undefined): boolean {
  return !!cell?.formula;
}

export { flattenRange, isError, isRange, unwrap };
